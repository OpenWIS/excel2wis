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

import sys
from lxml import etree

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
    tagsplit = str(tag).split(':')
    # tag name and prefix
    if len(tagsplit) == 2:
        prefix, tag_name = tagsplit
        el_list = parent.findall("{" + namespaces[prefix] + "}" + tag_name[:-3])
        new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name[:-3])
    # no prefix
    else:
        tag_name = str(tag)
        el_list = parent.findall(tag_name[:-3])
        new_element = parent.makeelement(tag_name[:-3])
    # A similar tag already exists in tree
    if el_list:
        el_list[-1].addnext(new_element)
    # New tag
    else:
        parent.append(new_element)
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

# Special case of free Keywords
# Add an ID attribute in MD_Keywords tag
def addAttributeIdKeywords(tree, xpath, attribute, att_id):
    xpath_list = xpath.split("/")[:]
    try:
        keyword_i = xpath_list.index('gmd:MD_Keywords')
        xpath_list = xpath_list[:keyword_i+1]
        xpath = "/".join(xpath_list)
        element = tree.xpath(xpath, namespaces=namespaces)[0] 
        element.attrib[attribute] = att_id
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
            sub_element = tree.xpath(xpath, namespaces=namespaces)[0]
        else:
            tagsplit = str(tag).split(':')
            if len(tagsplit) == 2:
                prefix, tag_name = tagsplit
                sub_element = etree.SubElement(parent, "{" + namespaces[prefix] + "}" + tag_name)
            else:
                tag_name = str(tag)
                sub_element = etree.SubElement(parent, tag_name)
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

def addOnlineResourceProtocol(tree, xpath_base):
    xpath_protocol = xpath_base + '/gmd:protocol/gco:CharacterString'
    addMetadataElement(tree, xpath_protocol, 'WWW:LINK-1.0-http--link')

# Find the multivalued element in xpath
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
        if tag == tag_list[-1] and value:
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

# Add dateType value and the associated codelist 
def addKeywordType(tree, xpath, type, code_list):
    # add the Keyword type written after Keyword
    value = type.split(':')[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:type', 'gmd:MD_KeywordTypeCode']
    xpath = "/".join(xpath_list)
    addMetadataElement(tree, xpath, value)
    addMetadataElement(tree, xpath, value, 'codeListValue')
    addMetadataElement(tree, xpath, code_list, 'codeList')



# Add resource format name
# and associated link, version
# Multiple format can be specified separated by ;
# name, version, specification are comma separated
def addResourceFormat(tree, xpath, value, urn):
    # Find multivalued tag xpath and list of afterwards tag to add for each occurrence
    name_tag_list, xpath = findMultiTagInXpath(tree, xpath)
    # Common tags to add once for each format
    base_tag_list = name_tag_list[:-2]
    # tag to add for specification and version
    specification_tag_list = ['gmd:specification', 'gco:CharacterString']
    version_tag_list = ['gmd:version', 'gco:CharacterString']
    # tag to add for name
    name_tag_list = name_tag_list[-2:]
    # Parse name, version and specification
    format_list = value.split(";")
    for format in format_list:
        val = format.split(",")
        try:
            name = val[0].strip()
            version = val[1].strip()
            specification = val[2].strip()
        except IndexError:
            sys.exit("%s Resource Format cell value is inconsistent with expected template" % urn)
        parent_xpath = xpath
        # Add common tags (and no value)
        addNewElementAndValue(tree, base_tag_list, '', parent_xpath)
        # Add specification (optionnal), version and name
        # order is important (the last one added is the first one in XML)
        base_xpath = xpath + "/" + "/".join(base_tag_list)
        if specification:
            addNewElementAndValue(tree, specification_tag_list, specification, base_xpath)
        addNewElementAndValue(tree, version_tag_list, version, base_xpath)
        addNewElementAndValue(tree, name_tag_list, name, base_xpath)

# Add GFNC file information
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
