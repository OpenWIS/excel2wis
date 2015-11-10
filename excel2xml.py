#!/usr/bin/env python
# -*- coding: utf-8 -*-

import copy
import sys
import xlrd
from lxml import etree
import xmltodict
import time


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
md_gene_value_col = 2
md_gene_xpath_col = 3
md_gene_codelist_col = 4
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

# Add several times the same tag (values comma separated)
def addMultiValue(tree, xpath, multivalue):
    xpath_list = xpath.split("/")[1:]
    xpath = ""
    # Rebuild of xpath to add missing tag
    for i, tag in enumerate(xpath_list):
        try:
            null, xpath = addMissingTags(tree, xpath, tag)
        except etree.XPathEvalError:
            break
    # MultiValue but no [] found in xpath
    if not tag.endswith('[]'):
        raise ValueError
    multi_tag_list = xpath_list[i:]
    multi_tag_list[0] = tag[:-2]
    for val in reversed(multivalue.split(',')):
        parent_xpath = xpath
        for tag in multi_tag_list:
            parent = tree.xpath(parent_xpath, namespaces=namespaces)[0]
            prefix, tag_name = str(tag).split(':')
            new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name)
            parent.insert(0, new_element)
            parent_xpath += "/" + tag
            if tag == multi_tag_list[-1]:
                new_element.text = val.strip()
    return parent_xpath

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
def addMetadataElement(tree, xpath, value, attribute='No'):
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
        el.attrib[attribute] = value
    return xpath

# Add tags which values are a concatenation that contains the urn
def concateValue(tree, value):
    # Unique Identifier (value for 1.3)
    urn = unicode(md_gene.cell_value(6, 2)).strip() + value
    # Location for online access
    nrow = 5 
    value = unicode(md_gene.cell_value(nrow, 2)).strip() + urn
    xpath = unicode(md_gene.cell_value(nrow, 3)).strip()
    addMetadataElement(tree, xpath, value)
    # Permanent link
    nrow = 4
    value = unicode(md_gene.cell_value(nrow, 2)).strip() + urn
    xpath = unicode(md_gene.cell_value(nrow, 3)).strip()
    addMetadataElement(tree, xpath, value)
    # Two linked tags are mandatory, cf. template (paragraph4)
    return urn

# Add dateType value and the codelist linked
def par1022(tree, xpath, type, code_list):
    # add the date type : creation, publication or revision
    value = type.split(':')[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:dateType', 'gmd:CI_DateTypeCode']
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
        if name.value == help_thesaurus:
            thes_i = i
    xpath_list = xpath.split('/')[:-2]
    xpath_th = "/".join(xpath_list)
    # Thesaurus informations in metadata XML
    # are different in a Format tag than in 
    # a keyword tag
    if 'gmd:MD_Format' in xpath:
        # Link
        xpath_th_link = xpath_th + '/gmd:specification/gco:CharacterString'
        addMetadataElement(tree, xpath_th_link,
                thesaurus_link[thes_i].value)
        # Version
        xpath_th_version = xpath_th + '/gmd:version/gco:CharacterString'
        addMetadataElement(tree, xpath_th_version,
                thesaurus_version[thes_i].value)
    elif 'gmd:MD_Keywords' in xpath:
        # Name
        xpath_th += '/gmd:thesaurusName/gmd:CI_Citation'
        xpath_th_name = xpath_th + "/gmd:title/gco:CharacterString"
        # Link
        addMetadataElement(tree, xpath_th_name,
            help_thesaurus + ' [' + thesaurus_link[thes_i].value + ']')
        # Date of revision
        date = thesaurus_date[thes_i].value
        if date:
            xpath_date = xpath_th + '/gmd:date/gmd:CI_Date/gmd:date/gco:Date'
            addMetadataElement(tree, xpath_date, date)
            xpath_datype = xpath_th + '/gmd:date/gmd:CI_Date/gmd:dateType/gmd:CI_DateTypeCode'
            datype = thesaurus_datype[thes_i].value
            addMetadataElement(tree, xpath_datype, datype)
            addMetadataElement(tree, xpath_datype, datype, 'codeListValue')
            datype_codelist = thesaurus_datype_codelist[thes_i].value
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
for row in range(md_gene_row_start, md_gene.nrows):
    value = unicode(md_gene.cell_value(row, md_gene_value_col)).strip()
    xpath = unicode(md_gene.cell_value(row, md_gene_xpath_col)).strip()
    code_list = unicode(md_gene.cell_value(row, md_gene_codelist_col)).strip()
    # empty Xpath
    if not xpath:
        empty_xpath_gene.append(row+1)
        continue
    try:
        addMetadataElement(common_tree, xpath, value)
        if code_list:
            addMetadataElement(common_tree, xpath, value, 'codeListValue')
            addMetadataElement(common_tree, xpath, code_list, 'codeList')
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
            if id == '1.3':
                uid = field_value
                # add tag values which are concatenation of MD generic and MD Fields elements
                urn = concateValue(tree, field_value)
                field_value = urn
            elif id == '6.3':
                # Value GTSPriority in Excel file does not validate
                field_value = 'GTSPriority' + field_value[9]
            
            # Add tags or attribute
            # Add several identical tags
            if multivalue == 'Yes':
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
