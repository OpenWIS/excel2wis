#!/usr/bin/env python
# -*- coding: utf-8 -*-

import copy
import sys
import xlrd
from lxml import etree
import xmltodict
import time
import datetime
import re


############################
# Excel file configuration #
############################
# Delta between MD Fields col and the linked
# Help row
# ID starts on the 2nd col of MD Fields
# and on the 5th row of Help
delta = 3
# MD Fields
fields_col_start = 1
fields_row_start = 6
# Help
type_col = 4
attribute_col = 5
thesaurus_col = 6
multivalue_col = 7
codelist_col = 8
att_id_col = 9
xpath_col = 10
# MD generic
md_gene_row_start = 3
md_gene_tag_col = 1
md_gene_value_col = 2
md_gene_xpath_col = 3
md_gene_codelist_col = 4
md_gene_attPrefix_col = 5
md_gene_attName_col = 6
md_gene_attValue_col = 7
# MD Thesaurus
thesaurus_col_start = 2
thesaurus_name_row = 2
thesaurus_link_row = 3
thesaurus_version_row = 4
thesaurus_datype_row = 5
thesaurus_datype_codelist_row = 7
thesaurus_date_row = 6

# Namespaces dict
namespaces = {'gmd': 'http://www.isotc211.org/2005/gmd',
              'gco': 'http://www.isotc211.org/2005/gco',
              'gfc': 'http://www.isotc211.org/2005/gfc',
              'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
              'xlink': 'http://www.w3.org/1999/xlink',
              'gml': 'http://www.opengis.net/gml/3.2',
              'gts': 'http://www.isotc211.org/2005/gts',
              'gmx': 'http://www.isotc211.org/2005/gmx'}


#####################################################
############# Adding DCPC tags ######################
#####################################################
def addDCPClinkage(urn, generic_dict):
    print "DCPC metadata - adding linkage"
    value_base = unicode(generic_dict['portal']['value']).strip() + '/openwis-user-portal/retrieve/'
    value = value_base + 'request/' + urn
    xpath = '/gmd:MD_Metadata/gmd:distributionInfo/gmd:MD_Distribution/gmd:transferOptions/gmd:MD_DigitalTransferOptions/gmd:onLine[]/gmd:CI_OnlineResource/gmd:linkage/gmd:URL'
    addMultiValueDCPC(tree, xpath, value, 'Request on DCPC')
    value = value_base + 'subscribe/' + urn
    addMultiValueDCPC(tree, xpath, value, 'Suscribe on DCPC')

# Adding linkage section for DCPC MD
def addMultiValueDCPC(tree, xpath, value, name):
    addMultiValue(tree, xpath, value)
    parent_xpath = '/gmd:MD_Metadata/gmd:distributionInfo/gmd:MD_Distribution/gmd:transferOptions/gmd:MD_DigitalTransferOptions/gmd:onLine[1]/gmd:CI_OnlineResource'
    addOnlineResourceProtocol(tree, parent_xpath)
    xpath_name = parent_xpath + '/gmd:name/gco:CharacterString'
    addMetadataElement(tree, xpath_name, name)

##### end of adding DCPC tags ######################

# Add an occurrence of an ordered tag missing from the template
# return xpath with the appropriate order (in case where an
# optional previous tag isn't filled)
def addMultipleElement(parent, xpath, tag):
    prefix, tag_name = str(tag).split(':')
    el_list = parent.findall("{" + namespaces[prefix] + "}" + tag_name[:-3])
    new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name[:-3])
    el_list[-1].addnext(new_element)
    new_element_index = len(el_list) + 1
    xpath = xpath[:-2] + str(new_element_index) + "]"
    return xpath

# Add attribute for generic metadata
def addAttribute(tree, xpath, prefix, name, value):
    prefix = prefix.split(',')
    name = name.split(',')
    value = value.split(',')
    if len(prefix) != len(name):
        print "MD_generic : please put as many attributes prefix as attributes names"
    for i, attName in enumerate(name):
        addMetadataElement(tree, xpath, value[i], attName, prefix[i])


# Extension of addAttributeIdKeywords
# Add an id attribute in the tag and with the prefix written in ID cell
# (not only in MD_Keywords)
#def addAttributeId(tree, xpath, att_id):
#    id_param = att_id.split(";")
#    att_value = id_param[0].strip()
#    where = id_param[1].strip()
#    prefix = ""
#    if len(id_param) > 2:
#        prefix = id_param[2].strip()
#    xpath_list = xpath.split("/")[:]
#    try:
#        keyword_i = xpath_list.index(where)
#        xpath_list = xpath_list[:keyword_i+1]
#        xpath = "/".join(xpath_list)
#        element = tree.xpath(xpath, namespaces=namespaces)[0] 
#        if prefix:
#            element.attrib["{" + namespaces[prefix] + "}" + "id"] = att_value
#        else:
#            element.attrib["id"] = att_value
#    except ValueError:
#        print "WARNING : ", where, " not found in XPATH"

# Special case of free Keywords
# Add an ID attribute in MD_Keywords tag
def addAttributeIdKeywords(tree, xpath, att_id):
    xpath_list = xpath.split("/")[:]
    try:
        keyword_i = xpath_list.index('gmd:MD_Keywords')
        xpath_list = xpath_list[:keyword_i+1]
        xpath = "/".join(xpath_list)
        element = tree.xpath(xpath, namespaces=namespaces)[0] 
        element.attrib['id'] = att_id
    except ValueError:
        print "WARNING : MD_Keywords not found in XPATH"

# Rebuild of XPATH to add missing tags
def addMissingTags(tree, xpath, tag):
    previous_xpath = xpath
    xpath += "/" + tag
    element = tree.xpath(xpath, namespaces=namespaces)
    sub_element = None
    # missing tag identified
    if len(element) == 0:
        # element under which the tag will be added
        parent = tree.xpath(previous_xpath, namespaces=namespaces)[0]
        # Add an occurrence of an ordered tag which is not in the template
        if xpath.endswith(']'):
            xpath = addMultipleElement(parent, xpath, tag)
            sub_element = tree.xpath(xpath, namespaces=namespaces)
        else:
            prefix, tag_name = str(tag).split(':')
            sub_element = etree.SubElement(parent, "{" + namespaces[prefix] + "}" + tag_name)
    return sub_element, xpath

# Add a single tag or attribute
# if the element already exists, its value is replaced
def addMetadataElement(tree, xpath, value, attribute='No', prefix='No'):
    element = tree.xpath(xpath, namespaces=namespaces)
    # Xpath found in the template
    if len(element) != 0:
        el = element[0]
    # Xpath not found in the template
    else:
        xpath_list = xpath.split("/")[1:]
        xpath = ""
        # Rebuild of xpath to add missing tag
        for i, tag in enumerate(xpath_list):
            el, xpath = addMissingTags(tree, xpath, tag)
    # Insert tag or attribute value
    if attribute == 'No':
        el.text = value
    else:
        if prefix == 'No':
            el.attrib[attribute] = value
        else:
            el.attrib["{" + namespaces[prefix] + "}" + attribute] = value
    return xpath

# Add tags which values are a concatenation that contains the urn
def concateValue(tree, value, generic_dict):
    # Unique Identifier
    urn = unicode(generic_dict['Unique identifier']['value']).strip() + value
    # Location for online access
    value = unicode(generic_dict['location (address) for on-line access']['value']).strip() + urn
    xpath = unicode(generic_dict['location (address) for on-line access']['xpath']).strip()
    addMetadataElement(tree, xpath, value)
    # URL permanent link
    value = unicode(generic_dict['permanent link']['value']).strip() + urn
    xpath = unicode(generic_dict['permanent link']['xpath']).strip()
    addMetadataElement(tree, xpath, value)
    # Two linked tags are mandatory, cf. template (paragraph4)
    return urn

def addOnlineResourceProtocol(tree, xpath_base):
    xpath_protocol = xpath_base + '/gmd:protocol/gco:CharacterString'
    addMetadataElement(tree, xpath_protocol, 'WWW:LINK-1.0-http--link')

# Find the multievaluated element in xpath
def findMultiTagInXpath(tree, xpath):
    xpath_list = xpath.split("/")[1:]
    xpath = ""
    # Add missing tags before [] and identify tag where [] is
    for i, tag in enumerate(xpath_list):
        try:
            null, xpath = addMissingTags(tree, xpath, tag)
        except etree.XPathEvalError:
            break
    # MultiValue set to Yes but no [] found in xpath
    if not tag.endswith('[]'):
        raise ValueError
    # List of tags to add several times
    # starting on the tag in which [] is found
    multi_tag_list = xpath_list[i:]
    # suppress [] in the mutlievaluated tag
    multi_tag_list[0] = tag[:-2]
    return multi_tag_list, xpath

# Create a new element and set its value (even if an element with the same xpath already exists)
def addNewElementAndValue(tree, tag_list, value, parent_xpath):
    for tag in tag_list:
        parent = tree.xpath(parent_xpath, namespaces=namespaces)[0]
        prefix, tag_name = str(tag).split(':')
        new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name)
        parent.insert(0, new_element)
        parent_xpath += "/" + tag
        if tag == tag_list[-1]:
            new_element.text = value.strip()
    return parent_xpath

# Add several times the same tag (values comma separated)
# new element are created
def addMultiValue(tree, xpath, multivalue):
    multi_tag_list, xpath = findMultiTagInXpath(tree, xpath)
    for val in reversed(multivalue.split(',')):
        parent_xpath = xpath
        parent_xpath = addNewElementAndValue(tree, multi_tag_list, val, parent_xpath)
    return parent_xpath

# Add link for resource locator (MD Fields)
# and associated protocol and name (3 elements for each link)
def addLink(tree, xpath, value):
    url_tag_list, xpath = findMultiTagInXpath(tree, xpath)
    base_tag_list = url_tag_list[:-2]
    url_tag_list_r = url_tag_list[-2:]
    # Add a new element (generic online resources are added and must not be erazed)
    name_tag_list = base_tag_list + ['gmd:name', 'gco:CharacterString']
    protocol_tag_list = ['gmd:protocol', 'gco:CharacterString']
    # parse name and URL
    # number of spaces around the colon can vary
    # "NAME URL" , "NAME URL" , "NAME URL"
    online_list = re.split("\xbb[\xa0 ]*,[\xa0 ]*\xab", value)
    for onliner in online_list:
        couple = re.search("\xab?(.*)[\xa0 ]*(https?://[^\xbb]*)", onliner.strip())
        or_name = couple.group(1).strip()
        or_URL = couple.group(2).strip()
        parent_xpath = xpath
        if or_name:
            addNewElementAndValue(tree, name_tag_list, or_name, parent_xpath)
        else:
            addNewElementAndValue(tree, name_tag_list, or_URL, parent_xpath)
        base_xpath = xpath + "/" + "/".join(base_tag_list)
        addNewElementAndValue(tree, protocol_tag_list, 'WWW:LINK-1.0-http--link', base_xpath)
        addNewElementAndValue(tree, url_tag_list_r, or_URL, base_xpath)

def addGFNC(tree, title, xpath, value):
    base = '/gmd:MD_Metadata/gmd:describes/gmx:MX_DataSet/'
    xpath_has = base + 'gmd:has'
    base = base + 'gmx:dataFile/gmx:MX_DataFile/'
    xpath_fileDescription = base + 'gmx:fileDescription/gco:CharacterString'
    xpath_fileType = base + 'gmx:fileType/gmx:MimeFileType'
    base = base + 'gmx:fileFormat/gmd:MD_Format/'
    xpath_fileFormat_name = base + 'gmd:name/gco:CharacterString'
    xpath_fileFormat_version = base + 'gmd:version/gco:CharacterString'
    addMetadataElement(tree, xpath_has, 'inapplicable', "{" + namespaces['gco'] + "}" + 'nilReason') 
    addMetadataElement(tree, xpath, value) 
    addMetadataElement(tree, xpath_fileDescription, title) 
    addMetadataElement(tree, xpath_fileType, 'application/octet-stream') 
    addMetadataElement(tree, xpath_fileType, 'application/octet-stream', 'type') 
    addMetadataElement(tree, xpath_fileFormat_name, 'BUFR') 
    addMetadataElement(tree, xpath_fileFormat_version, 'IV') 

# Add dateType value and the codelist linked
def par1022(tree, xpath, type, code_list):
    # add the date type : creation, publication or revision
    # written after "Date:"
    value = type.split(':')[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:dateType', 'gmd:CI_DateTypeCode']
    xpath = "/".join(xpath_list)
    addMetadataElement(tree, xpath, value)
    addMetadataElement(tree, xpath, value, 'codeListValue')
    addMetadataElement(tree, xpath, code_list, 'codeList')

# Add dateType value and the codelist linked
def addKeywordType(tree, xpath, type, code_list):
    # add the Keyword type written after Keyword
    value = type.split(':')[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:type', 'gmd:MD_KeywordTypeCode']
    xpath = "/".join(xpath_list)
    addMetadataElement(tree, xpath, value)
    addMetadataElement(tree, xpath, value, 'codeListValue')
    addMetadataElement(tree, xpath, code_list, 'codeList')

def addThesaurus(tree, xpath, help_thesaurus, thesaurus):
    # Name
    thesaurus_name = thesaurus.row(thesaurus_name_row)
    thesaurus_link = thesaurus.row(thesaurus_link_row)
    thesaurus_date = thesaurus.row(thesaurus_date_row)
    thesaurus_datype = thesaurus.row(thesaurus_datype_row)
    thesaurus_datype_codelist = thesaurus.row(thesaurus_datype_codelist_row)
    thesaurus_version = thesaurus.row(thesaurus_version_row)
    # Looking in the thesaurus sheet to find the col
    for i, name in enumerate(thesaurus_name):
        name_u = unicode(name.value).strip()
        if name_u == help_thesaurus:
            thes_i = i
    xpath_list = xpath.split('/')[:-2]
    xpath_th = "/".join(xpath_list)
    # Thesaurus informations in metadata XML
    # are different in a Format tag than in 
    # a keyword tag
    if 'gmd:MD_Format' in xpath:
        # Link
        xpath_th_link = xpath_th + '/gmd:specification/gco:CharacterString'
        link_u = unicode(thesaurus_link[thes_i].value).strip()
        addMetadataElement(tree, xpath_th_link, link_u)
        # Version
        xpath_th_version = xpath_th + '/gmd:version/gco:CharacterString'
        version_u = unicode(thesaurus_version[thes_i].value).strip()
        addMetadataElement(tree, xpath_th_version, version_u)
    elif 'gmd:MD_Keywords' in xpath:
        # Name
        xpath_th += '/gmd:thesaurusName/gmd:CI_Citation'
        xpath_th_name = xpath_th + "/gmd:title/gco:CharacterString"
        # Link
        # addMetadataElement(tree, xpath_th_name,
        #    help_thesaurus + ' [' + thesaurus_link[thes_i].value + ']')
        addMetadataElement(tree, xpath_th_name,
            help_thesaurus)
        # Date of revision
        date = unicode(thesaurus_date[thes_i].value).strip()
        if date:
            xpath_date = xpath_th + '/gmd:date/gmd:CI_Date/gmd:date/gco:Date'
            addMetadataElement(tree, xpath_date, date)
            xpath_datype = xpath_th + '/gmd:date/gmd:CI_Date/gmd:dateType/gmd:CI_DateTypeCode'
            datype = unicode(thesaurus_datype[thes_i].value).strip()
            addMetadataElement(tree, xpath_datype, datype)
            addMetadataElement(tree, xpath_datype, datype, 'codeListValue')
            datype_codelist = unicode(thesaurus_datype_codelist[thes_i].value).strip()
            addMetadataElement(tree, xpath_datype, datype_codelist, 'codeList')
    else:
        raise ValueError

###
# Excel file opening
###
try:
    excel_filename = sys.argv[1]
    workbook = xlrd.open_workbook(excel_filename)
except IndexError:
    sys.exit("Please put the name of the excel"
         + " file as the first argument")
except IOError:
    sys.exit("Excel filename %s not found" % excel_filename)
##
# Option [-openwis]
##
try:
    # Generate data-metadata csv file 
    openwis = sys.argv[2]
    if openwis == "-openwis":
        print "Creation of a new data metadata file"
        link_file = excel_filename.split('.')[0] + "_datalink.csv"
        open(link_file,'w').close()
    else:
        openwis = ""
except IndexError:
    openwis = ""

###
# Get sheets
###
md_fields = workbook.sheet_by_name('MD Fields')
help = workbook.sheet_by_name('Help')
md_gene = workbook.sheet_by_name('MD generic')
thesaurus = workbook.sheet_by_name('MD Thesaurus')
# ID on the 4th row of MD Fields
# and on the 2nd col of Help
field_id_list = md_fields.row(3)
field_mandatory_list = md_fields.row(4)
help_id_list = help.col(1)

###
# Track FATAL ERRORS
###
try:
    for i, id in enumerate(field_id_list):
        field_id = unicode(id.value).strip()
        help_id = unicode(help_id_list[i+delta].value).strip()
        # Check if MD Fields ID and Help ID match
        if field_id != help_id:
            raise Exception(
                    "ERROR : Paragraphs number in MD Fields sheet : "
                    + "%s doesn't match paragraphs" % field_id
                    + "number in Help sheet : %s" % help_id)
        # Check (non INSPIRE) mandatory fields
        # TODO INSPIRE : check mandatory fields for INSPIRE ?
        mandatory = unicode(field_mandatory_list[i].value).strip()
        if mandatory == 'Mandatory' and field_id not in ['8.1', '8.2']:
            for row in range(fields_row_start, md_fields.nrows):
                # mandatory field is not empty
                if not md_fields.cell_value(row, i): 
                    # when a cell value is 0, it must not be seen as an
                    # empty cell
                    if md_fields.cell_value(row, i) != 0:
                        raise Exception('ERROR - MD Fields sheet - ' +
                                        'Mandatory field %s' % field_id +
                                        ' is empty on row %s' % str(row+1))
                # xpath linked to mandatory field is not empty
                if not help.cell_value(i+delta, xpath_col):
                    raise Exception('ERROR - Help sheet - ' +
                                    'No XPATH linked to ' +
                                    'mandatory field %s' % help_id)

except IndexError:
    sys.exit("ERROR : There are more paragraphs in MD Fields sheet than in Help sheet")
except Exception as e:
    if e.args is None:
        sys.exit("ERROR : There is a mismatch between MD Fields and Help sheets")
    else:
        sys.exit(e.args[0])

###
# Read XML template
###
parser = etree.XMLParser(remove_blank_text=True)
common_tree = etree.parse("./template_WMO.xml", parser)

######################
# Add generic metadata 
# MD generic sheet
#######################

# Lists for WARN messages
empty_xpath_gene = []
error_gene = []
DCPC = False
###
# Create a dictionary for md_gene element
# used in specific metadata
###
generic_dict = {} 
for row in range(md_gene_row_start, md_gene.nrows):
    tag = unicode(md_gene.cell_value(row, md_gene_tag_col)).strip()
    value = unicode(md_gene.cell_value(row, md_gene_value_col)).strip()
    # Default datetime is utc timestamp
    if tag == 'Metadata date' and value == '':
        value = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    xpath = unicode(md_gene.cell_value(row, md_gene_xpath_col)).strip()
    code_list = unicode(md_gene.cell_value(row, md_gene_codelist_col)).strip()
    attPrefix = unicode(md_gene.cell_value(row, md_gene_attPrefix_col)).strip()
    attName = unicode(md_gene.cell_value(row, md_gene_attName_col)).strip()
    attValue = unicode(md_gene.cell_value(row, md_gene_attValue_col)).strip()
    tag_dict = {'value': value, 'xpath': xpath, 'codelist': code_list}
    generic_dict[tag] = tag_dict
    if not value: 
        #print "> empty MD generic field row : %s ignored" % str(int(row)+1) 
        continue    
    # empty Xpath
    if not xpath:
        empty_xpath_gene.append(row+1)
        continue
    try:
    # DCPC use case
        if tag.startswith('OpenWIS only') and value:
            DCPC = True
            print "DCPC metadata"
        addMetadataElement(common_tree, xpath, value)
        if tag.startswith('Resource locator') and tag.endswith('url'):
            xpath_base = "/".join(xpath.split('/')[:-2])
            addOnlineResourceProtocol(common_tree, xpath_base)
        if code_list:
            addMetadataElement(common_tree, xpath, value, 'codeListValue')
            addMetadataElement(common_tree, xpath, code_list, 'codeList')
        # Add attribute(s)
        if attName:
            addAttribute(common_tree, xpath, attPrefix, attName, attValue)
    except ValueError:
        error_gene.append(row+1)
        continue

# Write WARN messages for MD generic
if empty_xpath_gene or error_gene:
    print "\n--- WARN -------- MD generic ---"
    if empty_xpath_gene:
        print "elements on row(s) %s have no XPATH" % ", ".join([str(x) for x in empty_xpath_gene])
    if error_gene:
        print "elements on row(s) %s cannot be created, please check their xpath expression" % ", ".join([str(x) for x in error_gene])
    print "--------------------------------\n"

#######################
# Add specific metadata
#######################
# Iteration on MD Fields rows (one row = one metadata)
for row in range(fields_row_start, md_fields.nrows):
    tree = copy.deepcopy(common_tree)
    # Lists for WARN messages
    empty_xpath = []
    error = []
    for col in range(fields_col_start, md_fields.ncols):
        id = unicode(field_id_list[col].value).strip()  # element ID
        # TODO : translation
        if id.endswith('b'):
            # print "> translation %s ignored" % id
            continue
        # An optional empty field is not added
        mandatory = unicode(field_mandatory_list[col].value).strip()
        if mandatory == 'Optional':
            if not md_fields.cell_value(row, col): 
                # print "> empty optional field %s ignored" % id 
                continue    
        field_value = unicode(md_fields.cell_value(row, col)).strip()
        xpath = unicode(help.cell_value(col+delta, xpath_col)).strip()
        attribute = unicode(help.cell_value(col+delta, attribute_col)).strip()
        help_thesaurus = unicode(help.cell_value(col+delta, thesaurus_col)).strip()
        multivalue = unicode(help.cell_value(col+delta, multivalue_col)).strip()
        code_list = unicode(help.cell_value(col+delta, codelist_col)).strip()
        type = unicode(help.cell_value(col+delta, type_col)).strip()
        att_id = unicode(help.cell_value(col+delta, att_id_col)).strip()
        # empty Xpath
        if xpath == '':
            empty_xpath.append(id)
            continue

        try:
            #print "ID", id

            # Change of field value
            # Keep title for GFNC
            if xpath == '/gmd:MD_Metadata/gmd:identificationInfo/gmd:MD_DataIdentification/gmd:citation/gmd:CI_Citation/gmd:title/gco:CharacterString':
                title = field_value
            elif xpath == '/gmd:MD_Metadata/gmd:fileIdentifier/gco:CharacterString':
                uid = field_value
                # add tag values which are concatenation of MD generic and MD Fields elements
                urn = concateValue(tree, field_value, generic_dict)
                field_value = urn
                # DCPC
                if DCPC:
                    addDCPClinkage(urn, generic_dict)
            elif xpath == '/gmd:MD_Metadata/gmd:identificationInfo/gmd:MD_DataIdentification/gmd:resourceConstraints/gmd:MD_LegalConstraints/gmd:otherConstraints[2]/gco:CharacterString':
                # Value GTSPriority in Excel file does not validate
                field_value = 'GTSPriority' + field_value[9]
            elif xpath == '/gmd:MD_Metadata/gmd:describes/gmx:MX_DataSet/gmx:dataFile/gmx:MX_DataFile/gmx:fileName/gmx:FileName':
                # GFNC
                addGFNC(tree, title, xpath, field_value)
                # File to link data and metadata names [option -openwis]
                if openwis == "-openwis":
                    with open(link_file, 'a') as f:
                        f.write(urn + "," + field_value)

            # Online locator
            if xpath == '/gmd:MD_Metadata/gmd:distributionInfo/gmd:MD_Distribution/gmd:transferOptions/gmd:MD_DigitalTransferOptions/gmd:onLine[]/gmd:CI_OnlineResource/gmd:linkage/gmd:URL':
                 addLink(tree, xpath, field_value)
            # Add tags or attribute
            # Add several identical tags
            elif multivalue == 'Yes':
                xpath = addMultiValue(tree, xpath, field_value)
            # or add only one tag or attribute
            else:
                xpath = addMetadataElement(tree, xpath, field_value, attribute)
            
            # Add attribute ID in the MD_Keywords tag for free keywords
            if att_id:
                addAttributeIdKeywords(tree, xpath, att_id)

            # Add codelist
            # special case of Date (two fields must be filled : date and dateType)
            if type.startswith('Date:'):
                # add creation, publication or revision in dateType (paragraph 10.2.2)
                # the code_list is linked to de dateType
                par1022(tree, xpath, type, code_list)
            # normal case : addition of two attributes
            elif type.startswith('Keyword:'):
                # add KeywordType
                addKeywordType(tree, xpath, type, code_list)
            elif code_list:
                addMetadataElement(tree, xpath, field_value, 'codeListValue')
                addMetadataElement(tree, xpath, code_list, 'codeList')

            # Add thesaurus link, date and version
            if help_thesaurus:
                addThesaurus(tree, xpath, help_thesaurus, thesaurus)

        #except Exception as e:
        #    print id, e
        except ValueError:
            error.append(id)
            continue
        except etree.XPathEvalError:  # [] in xpath
            error.append(id)

    # Write an XML file for each metadata (row in MD Fields)
    metadata_row = row + 1
    string_xml = etree.tostring(tree, pretty_print=True, encoding='utf-8')
    # filename = "metadata_row" + str(metadata_row) + ".xml"
    date = time.strftime("%Y%m%d%H%M%S")
    filename = "MD_" + uid + "_" + date + ".xml"
    with open(filename, "wb") as fo:
        fo.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        fo.write(string_xml)

    print "\n##### File %s has been generated" % filename

    # Write WARN messages for MD Fields - Help for each row
    if empty_xpath or error:
        print "--- WARN -------- Fields row", metadata_row, "- Help ---"
        if empty_xpath:
            # not empty (optional) elements in MD Fields with no xpath linked in Help
            print "elements %s have no XPATH" % ", ".join(empty_xpath)
        if error:
            print "elements %s cannot be created, please check their xpath expression" % ", ".join(error)
    print "-----------------------------------------\n"
