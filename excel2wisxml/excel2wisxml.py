#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ---------------------------------------------------------------------
# Convert an excel template document to set of XML metadata
# compliant with WMO Core 1.3 profile
# Copyright (C) 2016  METEO FRANCE <gisc_support@meteo.fr>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
# ---------------------------------------------------------------------

import copy
import sys
import xlrd
from lxml import etree
import time
import datetime
import argparse
import codecs
import re
from excel2wisxmlutils import *
import os.path


#########################
# Add OpenWIS DCPC tags #
#########################
def addDCPClinkage(tree, urn, generic_dict):
    print "DCPC metadata - adding linkage"
    value_base = unicode(generic_dict['portal']['value']).strip() \
        + '/openwis-user-portal/retrieve/'
    value = value_base + 'request/' + urn
    xpath = '/gmd:MD_Metadata/gmd:distributionInfo/gmd:MD_Distribution/' + \
        'gmd:transferOptions/gmd:MD_DigitalTransferOptions/gmd:onLine[]' + \
        '/gmd:CI_OnlineResource/gmd:linkage/gmd:URL'
    addMultiValueDCPC(tree, xpath, value, 'Request on DCPC')
    value = value_base + 'subscribe/' + urn
    addMultiValueDCPC(tree, xpath, value, 'Subscribe on DCPC')


# Adding linkage section for DCPC MD
def addMultiValueDCPC(tree, xpath, value, name):
    addMultiValue(tree, xpath, value)
    parent_xpath = '/gmd:MD_Metadata/gmd:distributionInfo/' \
        + 'gmd:MD_Distribution/gmd:transferOptions/' \
        + 'gmd:MD_DigitalTransferOptions/gmd:onLine[1]/gmd:CI_OnlineResource'
    addOnlineResourceProtocol(tree, parent_xpath)
    xpath_name = parent_xpath + '/gmd:name/gco:CharacterString'
    addMetadataElement(tree, xpath_name, name)

##### end of add OpenWIS DCPC tags


# Add resource format information in GFNC section
def addResourceFormatGFNC(tree, name, version, mime, nb_filename):
    # check number of format and compare to number of filename
    nb_format = len(name)
    if nb_filename > nb_format:
        sys.exit("ERROR in Resource Format section : " +
                 "There must be either no value or at least as many " +
                 "Resource Format as filenames specified")
    xpath_dataFile = '/gmd:MD_Metadata/gmd:describes/gmx:MX_DataSet/' \
        + 'gmx:dataFile'
    for f in range(0, nb_filename):
        xpath_nb = "[%s]" % (f + 1)
        xpath_common = xpath_dataFile + xpath_nb + '/gmx:MX_DataFile/'
        xpath_mime = xpath_common + 'gmx:fileType/gmx:MimeFileType'
        addMetadataElement(tree, xpath_mime, mime[f])
        addMetadataElement(tree, xpath_mime, mime[f], 'type')
        xpath_format = xpath_common + 'gmx:fileFormat/gmd:MD_Format/'
        xpath_format_name = xpath_format + 'gmd:name/gco:CharacterString'
        xpath_format_version = xpath_format + 'gmd:version/gco:CharacterString'
        addMetadataElement(tree, xpath_format_name, name[f])
        addMetadataElement(tree, xpath_format_version, version[f])
        # Remove attribute nilReason for fileFormat and fileType
        element = tree.xpath(xpath_common + 'gmx:fileFormat',
                             namespaces=namespaces)[0]
        del element.attrib["{" + namespaces['gco'] + "}" + 'nilReason']
        element = tree.xpath(xpath_common + 'gmx:fileType',
                             namespaces=namespaces)[0]
        del element.attrib["{" + namespaces['gco'] + "}" + 'nilReason']


# Add GFNC file information
def addGFNC(tree, title, xpath, value):
    # Find multivalued tag xpath and list of afterwards tag
    # to add for each occurrence
    url_tag_list, xpath = findMultiTagInXpath(tree, xpath)
    # Common tags to add once for each resource locator
    base_tag_list = url_tag_list[:-2]
    # tag to add for name
    name_tag_list = ['gmx:fileName', 'gmx:FileName']
    common_format_tag_list = ['gmx:fileFormat']
    type_tag_list = ['gmx:fileType']
    description_tag_list = ['gmx:fileDescription', 'gco:CharacterString']
    parent_xpath = xpath
    # Add has tag
    xpath_has = parent_xpath + '/gmd:has'
    has_tag_list = ['gmd:has']
    # parse filename
    # number of spaces around the colon can vary
    # "filename1" , "filename2"
    online_list = re.split("[\"\xbb][\xa0 ]*,[\xa0 ]*[\"\xab]", value)
    for onliner in reversed(online_list):
        try:
            couple = re.search("[\"\xab]?([^\"\xbb]*)", onliner.strip())
            filename = couple.group(1).strip()
        except AttributeError:
            sys.exit("%s Free link cell value is inconsistent" % urn +
                     "with expected template")
        parent_xpath = xpath
        # Add new elements
        # (generic online resources added before and must not be erazed)
        # Add common tags (and no value)
        addNewElementAndValue(tree, base_tag_list, '', parent_xpath)
        # Add name (optionnal), protocol and url
        # order is important (the last one added is the first one in XML)
        base_xpath = xpath + "/" + "/".join(base_tag_list)
        format_parent = addNewElementAndValue(tree, common_format_tag_list,
                                              'inapplicable', base_xpath,
                                              "{" + namespaces['gco'] + "}" +
                                              'nilReason', isAttVal=False)
        addNewElementAndValue(tree, type_tag_list, 'inapplicable', base_xpath,
                              "{" + namespaces['gco'] + "}" + 'nilReason',
                              isAttVal=False)
        addNewElementAndValue(tree, description_tag_list, title, base_xpath)
        addNewElementAndValue(tree, name_tag_list, filename, base_xpath)
    addNewElementAndValue(tree, has_tag_list, 'inapplicable', xpath, "{" +
                          namespaces['gco'] + "}" + 'nilReason',
                          isAttVal=False)
    nb_filename = len(online_list)
    return nb_filename


# Add Temporal Extent Indeterminate Position
# in the attribute indeterminatePosition
# If the keyword "before" or "after" is identified
# the time position following it is added as a value
def addTemporalExtentIndeterminatePosition(tree, xpath, field_value):
    if field_value.startswith("before") or field_value.startswith("after"):
        temporalExtent = field_value.split(" ")
        addMetadataElement(tree, xpath, temporalExtent[0],
                           "indeterminatePosition")
        if len(temporalExtent) > 1:
            addMetadataElement(tree, xpath, temporalExtent[1])
    else:
        addMetadataElement(tree, xpath, field_value, "indeterminatePosition")


# Add link for resource locator
# and associated protocol and name (3 elements for each link)
# protocol is static, name and link are dynamic
# Multiple link can be specified (separated by ",")
# "name1 http://link1","name2 http://link2"
def addLink(tree, xpath, value, urn):
    # Find multivalued tag xpath and list of afterwards tag
    # to add for each occurrence
    url_tag_list, xpath = findMultiTagInXpath(tree, xpath)
    # Common tags to add once for each resource locator
    base_tag_list = url_tag_list[:-2]
    # tag to add for name and protocol
    name_tag_list = ['gmd:name', 'gco:CharacterString']
    protocol_tag_list = ['gmd:protocol', 'gco:CharacterString']
    # tag to add for URL
    url_tag_list = url_tag_list[-2:]
    # parse name and URL
    # number of spaces around the colon can vary
    # "NAME URL" , "NAME URL" , "NAME URL"
    online_list = re.split("[\"\xbb][\xa0 ]*,[\xa0 ]*[\"\xab]", value)
    for onliner in online_list:
        try:
            couple = re.search("[\"\xab]?(.*)[\xa0 ]*(https?://[^\"\xbb]*)",
                               onliner.strip())
            or_name = couple.group(1).strip()
            or_URL = couple.group(2).strip()
        except AttributeError:
            sys.exit("%s Free link cell value is inconsistent with" % urn +
                     "expected template")
        parent_xpath = xpath
        # Add new elements
        # (generic online resources added before and must not be erazed)
        # Add common tags (and no value)
        addNewElementAndValue(tree, base_tag_list, '', parent_xpath)
        # Add name (optionnal), protocol and url
        # order is important (the last one added is the first one in XML)
        base_xpath = xpath + "/" + "/".join(base_tag_list)
        # if name is not defined, default value is the URL
        if or_name:
            addNewElementAndValue(tree, name_tag_list, or_name, base_xpath)
        else:
            addNewElementAndValue(tree, name_tag_list, or_URL, base_xpath)
        addNewElementAndValue(tree, protocol_tag_list,
                              'WWW:LINK-1.0-http--link', base_xpath)
        addNewElementAndValue(tree, url_tag_list, or_URL, base_xpath)


# Add thesaurus information
def addThesaurus(tree, xpath, help_thesaurus, thesaurus, thesaurus_rows):
    thesaurus_name = thesaurus.row(thesaurus_rows["name"])
    thesaurus_link = thesaurus.row(thesaurus_rows["link"])
    thesaurus_date = thesaurus.row(thesaurus_rows["date"])
    thesaurus_datype = thesaurus.row(thesaurus_rows["datetype"])
    thesaurus_datype_codelist = thesaurus.row(
        thesaurus_rows["datetypecodelist"])
    thesaurus_version = thesaurus.row(thesaurus_rows["version"])
    # Looking in the thesaurus sheet to find the col
    for i, name in enumerate(thesaurus_name):
        name_u = unicode(name.value).strip()
        if name_u == help_thesaurus:
            thes_i = i
    xpath_list = xpath.split('/')[:-2]
    xpath_th = "/".join(xpath_list)
    if 'gmd:MD_Keywords' in xpath:
        # Name
        xpath_th += '/gmd:thesaurusName/gmd:CI_Citation'
        xpath_th_name = xpath_th + "/gmd:title/gco:CharacterString"
        # Link
        # addMetadataElement(tree, xpath_th_name,
        #    help_thesaurus + ' [' + thesaurus_link[thes_i].value + ']')
        addMetadataElement(tree, xpath_th_name, help_thesaurus)
        # Date of revision
        date = unicode(thesaurus_date[thes_i].value).strip()
        if date:
            xpath_date = xpath_th + '/gmd:date/gmd:CI_Date/gmd:date/gco:Date'
            addMetadataElement(tree, xpath_date, date)
            xpath_datype = xpath_th + \
                '/gmd:date/gmd:CI_Date/gmd:dateType/gmd:CI_DateTypeCode'
            datype = unicode(thesaurus_datype[thes_i].value).strip()
            addMetadataElement(tree, xpath_datype, datype)
            addMetadataElement(tree, xpath_datype, datype, 'codeListValue')
            datype_codelist = unicode(
                thesaurus_datype_codelist[thes_i].value).strip()
            addMetadataElement(tree, xpath_datype, datype_codelist, 'codeList')
    else:
        raise ValueError


# Add tags for which values are a concatenation that contains the urn
def concateValue(tree, value, generic_dict):
    # Unique Identifier
    urn = unicode(generic_dict['Unique identifier']['value']).strip() + value
    # Location for online access
    value = unicode(generic_dict[
        'location (address) for on-line access']['value']).strip() + urn
    xpath = unicode(generic_dict[
        'location (address) for on-line access']['xpath']).strip()
    addMetadataElement(tree, xpath, value)
    # URL permanent link
    value = unicode(generic_dict['permanent link']['value']).strip() + urn
    xpath = unicode(generic_dict['permanent link']['xpath']).strip()
    addMetadataElement(tree, xpath, value)
    # Two linked tags are mandatory, cf. template (paragraph4)
    return urn


# Add elements for translations
def addTranslation(tree, xpath, translation_value, secondLanguage):
    # add translation attribute in parent element
    translation_location = xpath.split('/')[-2]
    addAttribute(tree, xpath, 'xsi:type', 'gmd:PT_FreeText_PropertyType',
                 translation_location)
    # add translation element as a sibling
    translation_xpath = '/'.join(xpath.split('/')[:-1]) + \
        "/gmd:PT_FreeText/gmd:textGroup/gmd:LocalisedCharacterString"
    addMetadataElement(tree, translation_xpath, translation_value)
    addMetadataElement(tree, translation_xpath, '#locale-' +
                       secondLanguage, 'locale')


# Add elements for multiValue translations
def addMultiValueTranslation(tree, xpath, translation_value, secondLanguage):
    element = tree.xpath(xpath, namespaces=namespaces)
    for i, val in enumerate(translation_value.split(",")):
        el = element[i]
        # add translation attribute in parent element
        parent = el.getparent()
        parent.attrib[
            "{" + namespaces["xsi"] + "}type"] = "gmd:PT_FreeText_PropertyType"
        # add translation element as a sibling
        sibling = etree.Element("{" + namespaces["gmd"] + "}PT_FreeText")
        el.addnext(sibling)
        sibling = el.getnext()
        textGroup = etree.Element("{" + namespaces["gmd"] + "}textGroup")
        sibling.insert(0, textGroup)
        localised = etree.Element("{" + namespaces["gmd"] +
                                  "}LocalisedCharacterString")
        localised.text = val.strip()
        localised.attrib["locale"] = "#locale-" + secondLanguage
        textGroup.insert(0, localised)


def addLocaleInfo(tree, xpath, secondLanguage):
    xpath_base = "/".join(xpath.split("/")[:-2])
    xpath_encoding = xpath_base + \
        "/gmd:characterEncoding/gmd:MD_CharacterSetCode"
    addMetadataElement(tree, xpath_base, "locale-" + secondLanguage, "id")
    addMetadataElement(tree, xpath_encoding, "utf-8")
    addMetadataElement(tree, xpath_encoding,
                       "resources/Codelist/gmxcodelists.xml" +
                       "#MD_CharacterSetCode", "codeList")
    addMetadataElement(tree, xpath_encoding, "utf-8", 'codeListValue')


# Create metadata from excel file
def excel2wisxml(excel_filename, MFopenwis=False):

    base_path = os.path.dirname(__file__)

    SCRIPT_VERSION = "4"
    EXCEL_FIRST_COMPATIBLE_VERSION = "3.3"

# Excel file location
    excel_path = os.path.dirname(excel_filename)
# Excel file name
    excel_name = os.path.basename(excel_filename)

###
# Print license information
###
    print "--------------------------------------------------------------"
    print "excel2wisxml  Copyright (C) 2016  METEO FRANCE"
    print "This program comes with ABSOLUTELY NO WARRANTY."
    print "This is free software, and you are welcome to redistribute it"
    print "under certain conditions."
    print "--------------------------------------------------------------"

###
# Excel file opening
###
    try:
        workbook = xlrd.open_workbook(excel_filename)
    except IOError:
        sys.exit("Excel filename %s not found" % excel_filename)

    if MFopenwis:
        date = time.strftime("%Y%m%d%H%M%S")
        link_file = excel_filename[:-4] + "_" + date + ".csv"
        codecs.open(link_file, 'w', 'utf-8').close()

###
# Get sheets
###
    md_fields = workbook.sheet_by_name('MD Fields')
    help = workbook.sheet_by_name('Help')
    md_gene = workbook.sheet_by_name('MD generic')
    thesaurus = workbook.sheet_by_name('MD Thesaurus')
# Get translation sheets
    md_fields_translation = workbook.sheet_by_name('MD Fields Translate')

##################################
# Excel file shape configuration #
##################################

# Delta between MD Fields col
# and the linked Help row
# ID starts on the 2nd col of MD Fields
# and on the 5th row of Help
    delta = 3

# MD Fields column and row start
    fields_col_start = 1
    fields_row_start = 6
    fields_row_mandatory = 4
# Section number row
    fields_row_section = 3

# Help
# Associate columns and headers
    help_header = help.row(2)
    for i, head in enumerate(help_header):
        head = head.value.strip().lower()
        if head == 'type':
            type_col = i
        elif head == 'attribute name':
            att_name_col = i
        elif head == 'attribute location':
            att_loc_col = i
        elif head == 'thesaurus name':
            thesaurus_col = i
        elif head == 'multi value':
            multivalue_col = i
        elif head == 'codelist':
            codelist_col = i
        elif head == 'attribute value':
            att_val_col = i
        elif head == 'xpath':
            xpath_col = i
        elif head == 'section':
            section_col = i

# MD generic row start
    md_gene_row_start = 3
# Associate columns and headers
    md_gene_header = md_gene.row(2)
    for i, head in enumerate(md_gene_header):
        head = head.value.strip().lower()
        if head == 'tag':
            md_gene_tag_col = i
        elif head == 'value':
            md_gene_value_col = i
        elif head == 'value translation (optional)':
            md_gene_translation_value_col = i
        elif head == 'xpath':
            md_gene_xpath_col = i
        elif head == 'codelist':
            md_gene_codelist_col = i
        elif head == 'attribut: location':
            md_gene_attLocation_col = i
        elif head == 'attribut: name':
            md_gene_attName_col = i
        elif head == 'attribut: value':
            md_gene_attValue_col = i

# MD Thesaurus column start
    thesaurus_col_start = 2
# Associate columns and headers
    thesaurus_header = thesaurus.col(1)
    for i, head in enumerate(thesaurus_header):
        head = head.value.strip().lower()
        if head == 'name':
            thesaurus_name_row = i
        elif head == 'link':
            thesaurus_link_row = i
        elif head == 'version':
            thesaurus_version_row = i
        elif head == 'date type':
            thesaurus_datype_row = i
        elif head == 'date':
            thesaurus_date_row = i
        elif head == 'date type codelist':
            thesaurus_datype_codelist_row = i

    thesaurus_rows = {"name": thesaurus_name_row,
                      "link": thesaurus_link_row,
                      "version": thesaurus_version_row,
                      "datetype": thesaurus_datype_row,
                      "date": thesaurus_date_row,
                      "datetypecodelist": thesaurus_datype_codelist_row}

# Get sections numbers lists for Help and MD Fields
    field_id_list = md_fields.row(fields_row_section)
    help_id_list = help.col(section_col)

# Get mandatory list for MD Fields
    field_mandatory_list = md_fields.row(fields_row_mandatory)

### End of excel file shape configuration

###
# Track FATAL ERRORS
###
    try:
        for i, id in enumerate(field_id_list):
            field_id = unicode(id.value).strip()
            help_id = unicode(help_id_list[i + delta].value).strip()
            # Check if MD Fields ID and Help ID match
            if field_id != help_id:
                raise Exception(
                    "ERROR : Paragraphs number in MD Fields sheet : " +
                    "%s doesn't match paragraphs" % field_id +
                    "number in Help sheet : %s" % help_id)
            # Check mandatory fields
            mandatory = unicode(field_mandatory_list[i].value).strip()
            if mandatory == 'Mandatory':
                for row in range(fields_row_start, md_fields.nrows):
                    # mandatory field is not empty
                    if not md_fields.cell_value(row, i):
                        # when a cell value is 0, it must not be seen as an
                        # empty cell
                        if md_fields.cell_value(row, i) != 0:
                            raise Exception('ERROR - MD Fields sheet - ' +
                                            'Mandatory field %s ' % field_id +
                                            'is empty on row %s' %
                                            str(row + 1))
                    # xpath linked to mandatory field is not empty
                    if not help.cell_value(i + delta, xpath_col):
                        raise Exception('ERROR - Help sheet - ' +
                                        'No XPATH linked to ' +
                                        'mandatory field %s' % help_id)

    except IndexError:
        sys.exit("ERROR : There are more paragraphs in MD Fields sheet" +
                 "than in Help sheet")
    except Exception as e:
        if e.args is None:
            sys.exit("ERROR : " +
                     "There is a mismatch between MD Fields and Help sheets")
        else:
            sys.exit(e.args[0])

###
# Read XML template
###
    parser = etree.XMLParser(remove_blank_text=True)
    common_tree = etree.parse(base_path +
                              "/templates/excel2wisxml_template.xml", parser)

#######################
# Add generic metadata
# MD generic sheet
#######################

# Option --MFopenwis ERROR message
    option_error = False
# Lists for WARN messages
    empty_xpath_gene = []
    error_gene = []
    DCPC = False
    translation = False
###
# Create a dictionary for md_gene element
# used in specific metadata
###
    generic_dict = {}
    for row in range(md_gene_row_start, md_gene.nrows):
        tag = unicode(md_gene.cell_value(row, md_gene_tag_col)).strip()
        value = unicode(md_gene.cell_value(row, md_gene_value_col)).strip()
        translation_value = unicode(
            md_gene.cell_value(row, md_gene_translation_value_col)).strip()
        xpath = unicode(md_gene.cell_value(row, md_gene_xpath_col)).strip()
        code_list = unicode(
            md_gene.cell_value(row, md_gene_codelist_col)).strip()
        attLocation = unicode(
            md_gene.cell_value(row, md_gene_attLocation_col)).strip()
        attName = unicode(md_gene.cell_value(row, md_gene_attName_col)).strip()
        attValue = unicode(
            md_gene.cell_value(row, md_gene_attValue_col)).strip()
        tag_dict = {'value': value, 'xpath': xpath, 'codelist': code_list}
        generic_dict[tag] = tag_dict
        if not md_gene.cell_type(row, md_gene_value_col):
            # Default datetime is utc timestamp
            if tag == 'Metadata date':
                value = datetime.datetime.utcnow().strftime(
                    "%Y-%m-%dT%H:%M:%SZ")
            else:
                # print "> empty MD generic field row:
                # %s ignored" % str(int(row)+1)
                continue
        # empty Xpath
        if not xpath:
            # Get excel version
            if tag.startswith('ExcelVersion'):
                excel_version = value
                # Check if excel version is compatible with script version
                int_excel_version = int(excel_version.replace(".", ""))
                int_excel_compatible_version = int(
                    EXCEL_FIRST_COMPATIBLE_VERSION.replace(".", ""))
                if int_excel_version < int_excel_compatible_version:
                    sys.exit("Metadata excel file version %s is not " %
                             excel_version + "compatible with this script. " +
                             "Please use version %s. " %
                             EXCEL_FIRST_COMPATIBLE_VERSION)
            else:
                empty_xpath_gene.append(row + 1)
            continue
        try:
            # identify DCPC use case
            if tag.startswith('OpenWIS only') and value:
                DCPC = True
            addMetadataElement(common_tree, xpath, value)
        # identify translation use case
        # save second language
        # and add elements
            if tag.startswith('Metadata second language') and value:
                translation = True
                secondLanguage = value
                addLocaleInfo(common_tree, xpath, secondLanguage)
            if tag.startswith('Resource locator') and tag.endswith('url'):
                xpath_base = "/".join(xpath.split('/')[:-2])
                addOnlineResourceProtocol(common_tree, xpath_base)
            if code_list:
                addMetadataElement(common_tree, xpath, value, 'codeListValue')
                addMetadataElement(common_tree, xpath, code_list, 'codeList')
            # Add attribute(s)
            if attName:
                addAttribute(common_tree, xpath, attName,
                             attValue, attLocation)
            # Add translations
            if translation and translation_value:
                addTranslation(common_tree, xpath,
                               translation_value, secondLanguage)
        except SystemExit as e:
            # sys.exit() generates a SystemExit exception
            print "ERROR section", id, e
            sys.exit()
        except:
            sys.exit("MD generic tag %s\n\terror in xpath %s" % (tag, xpath))

# Write WARN messages for MD generic
    if empty_xpath_gene:
        print "\n--- WARN -------- MD generic ---"
        print "elements on row(s) %s have no XPATH" % ", ".join(
            [str(x) for x in empty_xpath_gene])
        print "--------------------------------\n"

# Print version number in CSV file
    if MFopenwis:
        with codecs.open(link_file, 'a', 'utf-8') as f:
            f.write("EXCEL_VERSION=%s SCRIPT_VERSION=%s" % (excel_version,
                                                            SCRIPT_VERSION))

#######################
# Add specific metadata
#######################
# Iteration on MD Fields rows (one row = one metadata)
    for row in range(fields_row_start, md_fields.nrows):
        emptyDescriptiveKeywords = []
        # number of filenames specified (if not null
        # resource format information is added)
        nb_filename = 0
        tree = copy.deepcopy(common_tree)
        # Lists for WARN messages
        empty_xpath = []
        gfnc = ""

        # Put translations in a dictionary for MD Fields current row
        translation_fields = {}
        if translation:
            for col in range(fields_col_start, md_fields_translation.ncols):
                id = unicode(md_fields_translation.cell_value(
                    fields_row_section, col)).strip()
                if id.startswith('lg:'):
                    id = id[3:]
                    value = unicode(md_fields_translation.cell_value(
                        row, col)).strip()
                    translation_fields[id] = value

        for col in range(fields_col_start, md_fields.ncols):
            id = unicode(field_id_list[col].value).strip()  # element ID
            # An optional empty field is not added
            mandatory = unicode(field_mandatory_list[col].value).strip()
            xpath = unicode(help.cell_value(col + delta, xpath_col)).strip()
            if mandatory == 'Optional':
                if not md_fields.cell_type(row, col):
                    if 'descriptiveKeywords' in xpath:
                        xpath = xpath.split("/gmd:MD_Keywords")[0]
                        emptyDescriptiveKeywords.append(xpath)
                        # list of xpath of empty descriptiveKeywords
                    continue
            field_value = unicode(md_fields.cell_value(row, col)).strip()
            field_value_lower = field_value.lower()
            att_name = unicode(help.cell_value(col + delta,
                                               att_name_col)).strip()
            att_location = unicode(help.cell_value(col + delta,
                                                   att_loc_col)).strip()
            help_thesaurus = unicode(help.cell_value(col + delta,
                                                     thesaurus_col)).strip()
            multivalue = unicode(help.cell_value(col + delta,
                                                 multivalue_col)).strip()
            code_list = unicode(help.cell_value(col + delta,
                                                codelist_col)).strip()
            element_type = unicode(help.cell_value(col + delta,
                                                   type_col)).strip()
            att_val = unicode(help.cell_value(col + delta,
                                              att_val_col)).strip()
            att_val_exception = att_val.startswith('MD_Fields')
            # empty Xpath
            if xpath == '':
                empty_xpath.append(id)
                continue

            try:
                # Change value or keep it in memory
                if xpath == '/gmd:MD_Metadata/gmd:identificationInfo/' + \
                        'gmd:MD_DataIdentification/gmd:citation/' + \
                        'gmd:CI_Citation/gmd:title/gco:CharacterString':
                    # Keep title for GFNC
                    title = field_value
                elif (xpath == '/gmd:MD_Metadata/gmd:fileIdentifier/' +
                               'gco:CharacterString'):
                    uid = field_value
                    # add tag values which are concatenation of MD generic
                    # and MD Fields elements
                    # Keep URN value
                    urn = concateValue(tree, field_value, generic_dict)
                    field_value = urn
                    # DCPC
                    if DCPC:
                        print "OpenWIS DCPC metadata"
                        addDCPClinkage(tree, urn, generic_dict)
                elif (xpath == '/gmd:MD_Metadata/gmd:identificationInfo/' +
                        'gmd:MD_DataIdentification/gmd:resourceConstraints/' +
                        'gmd:MD_LegalConstraints/gmd:otherConstraints[2]/' +
                        'gco:CharacterString'):
                    # Value GTSPriority in Excel file does not validate
                    field_value = 'GTSPriority' + field_value[9]

                # Specific processing
                # (cell value is not added exactly at XPATH location)
                # Free links
                # Temporal Extent
                temporal_extent_xpath = [
                    '/gmd:MD_Metadata/gmd:identificationInfo/' +
                    'gmd:MD_DataIdentification/gmd:extent/gmd:EX_Extent/' +
                    'gmd:temporalElement/gmd:EX_TemporalExtent/gmd:extent/' +
                    'gml:TimePeriod/gml:beginPosition',
                    '/gmd:MD_Metadata/gmd:identificationInfo/' +
                    'gmd:MD_DataIdentification/gmd:extent/gmd:EX_Extent/' +
                    'gmd:temporalElement/gmd:EX_TemporalExtent/gmd:extent/' +
                    'gml:TimePeriod/gml:endPosition']
                temporal_extent_attribut = ["now", "unknown",
                                            "after", "before"]
                if xpath in temporal_extent_xpath and (
                        field_value_lower in temporal_extent_attribut or
                        "before" in field_value_lower or
                        "after" in field_value_lower):
                    # Put the value in attribute indeterminatePosition
                    addTemporalExtentIndeterminatePosition(tree, xpath,
                                                           field_value_lower)
                elif (xpath == '/gmd:MD_Metadata/gmd:distributionInfo/' +
                        'gmd:MD_Distribution/gmd:transferOptions/' +
                        'gmd:MD_DigitalTransferOptions/gmd:onLine[]/' +
                        'gmd:CI_OnlineResource/gmd:linkage/gmd:URL'):
                    addLink(tree, xpath, field_value, urn)
                # Resource Format
                elif (xpath == '/gmd:MD_Metadata/gmd:distributionInfo/' +
                      'gmd:MD_Distribution/gmd:distributionFormat[]/' +
                      'gmd:MD_Format/gmd:name/gco:CharacterString'):
                    name_list, version_list, spec_list, mime_list = \
                        addResourceFormat(tree, xpath, field_value, urn)
                    if nb_filename:
                        addResourceFormatGFNC(tree, name_list, version_list,
                                              mime_list, nb_filename)
                elif (xpath == '/gmd:MD_Metadata/gmd:describes/' +
                      'gmx:MX_DataSet/gmx:dataFile[]/gmx:MX_DataFile' +
                      '/gmx:fileName/gmx:FileName'):
                    # GFNC
                    nb_filename = addGFNC(tree, title, xpath, field_value)
                    if MFopenwis:
                        gfnc = field_value
                # Add tags or attribute
                # Add several identical tags
                elif multivalue == 'Yes':
                    xpath = addMultiValue(tree, xpath, field_value)
                # or add only one tag or attribute
                elif att_val_exception:
                    # Attribute read in MD_Fields sheet
                    xpath = addMetadataElement(tree, xpath,
                                               field_value, att_name)

                # Regular processing
                # (cell value is added at the XPATH specified in Help sheet)
                else:
                    xpath = addMetadataElement(tree, xpath, field_value)

                # Add elements in addition to one just added
                # Add attribute
                if att_name != 'No' and not att_val_exception:
                    addAttribute(tree, xpath, att_name, att_val, att_location)

                # special case of Date
                if element_type.startswith('Date:'):
                    # add creation, publication
                    # or revision in dateType (paragraph 10.2.2)
                    # the code_list is linked to the dateType
                    addDateType(tree, xpath, element_type, code_list)
                # normal case : addition of two attributes
                elif element_type.startswith('Keyword:'):
                    # add KeywordType
                    addKeywordType(tree, xpath, element_type, code_list)
                # Add codelist
                elif code_list:
                    addMetadataElement(tree, xpath,
                                       field_value, 'codeListValue')
                    addMetadataElement(tree, xpath, code_list, 'codeList')

                # Add thesaurus link, date and version
                if help_thesaurus:
                    addThesaurus(tree, xpath,
                                 help_thesaurus, thesaurus, thesaurus_rows)

                # Add translations
                if translation and id in translation_fields:
                    if multivalue == 'Yes':
                        addMultiValueTranslation(tree, xpath,
                                                 translation_fields[id],
                                                 secondLanguage)
                    else:
                        addTranslation(tree, xpath,
                                       translation_fields[id], secondLanguage)

            #except Exception as e:
            #    print id, e
            except SystemExit as e:
                # sys.exit() generates a SystemExit exception
                print "ERROR row", row + 1, "section", id, e
                sys.exit()
            except:
                sys.exit("MD Fields section %s\n\terror in xpath %s" %
                         (id, xpath))

        # Remove empty descriptiveKeywords
        if emptyDescriptiveKeywords:
            # Sort xpath according to descriptiveKeywords index
            # top down to remove the highest index first
            # (otherwise xpath index changes before its deletion)
            index_emptyDK = [int(re.search(
                "gmd:descriptiveKeywords\[([1-9]*)\]",
                dk).group(1)) for dk in emptyDescriptiveKeywords]
            index_emptyDK = sorted(index_emptyDK, reverse=True)
            emptyDescriptiveKeywords = \
                [dk[:-2] + str(i) + "]" for dk, i in zip(
                    emptyDescriptiveKeywords, index_emptyDK)]
        for xpath in emptyDescriptiveKeywords:
            try:
                element = tree.xpath(xpath, namespaces=namespaces)[0]
                element.getparent().remove(element)
            except:
                pass

        # Write an XML file for each metadata (row in MD Fields)
        metadata_row = row + 1
        string_xml = etree.tostring(tree, pretty_print=True, encoding='utf-8')
        date = time.strftime("%Y%m%d%H%M%S")
        filename = os.path.join(excel_path, "MD_" + uid + "_" + date + ".xml")
        with open(filename, "wb") as fo:
            fo.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            fo.write('<!-- Metadata generated with ' +
                     'Metadata-guide-record.xls version %s ' % excel_version +
                     'and excel2wisxml.py version %s -->\n' % SCRIPT_VERSION)
            fo.write(string_xml)

        if MFopenwis:
            if gfnc:
                with codecs.open(link_file, 'a', 'utf-8') as f:
                    f.write("\n\"" + urn + "\" ; \"" + gfnc + "\"")
            else:
                option_error = True


        print "\n##### File %s has been generated" % filename

        # Write WARN messages for MD Fields - Help for each row
        if empty_xpath:
            print "--- WARN -------- Fields row", metadata_row, "- Help ---"
            # not empty (optional) elements in MD Fields
            # with no xpath linked in Help
            print "elements %s have no XPATH" % ", ".join(empty_xpath)
        print "-----------------------------------------\n"

    if MFopenwis:
        if option_error:
            print "WARNING --MFopenwis"
            print "CSV file has not been generated"
            print "MD Fields file name section must be filled" + \
                "for each metadata in excel file to generate CSV file\n"
        else:
            print "CSV file %s has been generated\n" % link_file


def main():
    ###
    # Script help configuration
    # and arguments retrieval
    ###
    parser = argparse.ArgumentParser(
        description='Create a WMO Core Profile 1.3 XML file ' +
        'from an excel file.')
    parser.add_argument(
        'filename', metavar='filename', type=str, nargs=1,
        help='Excel file name containing metadata information')
    parser.add_argument(
        '--MFopenwis', action='store_true',
        help='option to generate a CSV file containing metadata URNs ' +
        'and associated data file name')
    args = parser.parse_args()
    excel_filename = args.filename[0]
    MFopenwis = args.MFopenwis
# Call main function to create xml metadata file
    excel2wisxml(excel_filename, MFopenwis)


def createExcel():
    ###
    # Copy excel template in the current directory
    ###
    base_path = os.path.dirname(__file__)
    date = time.strftime("%Y%m%d%H%M%S")
    command = "cp %s/templates/Metadata-guide-record.xls ./Metadata-guide-record-%s.xls" % (base_path, date)
    os.system(command)
