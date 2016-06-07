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
        tag_name = tag_name.split("[")[0]
        el_list = parent.findall("{" + namespaces[prefix] + "}" + tag_name)
        new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name)
    # no prefix
    else:
        tag_name = str(tag)
        tag_name = tag_name.split("[")[0]
        el_list = parent.findall(tag_name.split("[")[0])
        new_element = parent.makeelement(tag_name)
    # A similar tag already exists in tree
    if el_list:
        el_list[-1].addnext(new_element)
    # New tag
    else:
        parent.append(new_element)
    new_element_index = len(el_list) + 1
    xpath_list = xpath.split('/')
    xpath_list[-1] = xpath_list[-1].split("[")[0] + "[" + str(new_element_index) + "]"
    xpath = "/".join(xpath_list)
    return xpath

# Add attribute for generic metadata
#def addAttribute(tree, xpath, name, value):
#    name = name.split(',')
#    value = value.split(',')
#    for i, attName in enumerate(name):
#        # is there a prefix
#        tagsplit = str(attName).split(':')
#        if len(tagsplit) == 2:
#            prefix, att_name = tagsplit
#            addMetadataElement(tree, xpath, value[i], att_name, prefix)
#        else:
#            att_name = str(attName)
#            addMetadataElement(tree, xpath, value[i], att_name)

# Special case of free Keywords
# Add an ID attribute in MD_Keywords tag
def addAttribute(tree, xpath, att_name, att_val, att_loc):
    att_name = att_name.split(',')
    att_val = att_val.split(',')
    if att_loc:
        att_loc = att_loc.split(',')
    # If Attribute Location col is not filled at all
    # it is replaced by an empty list
    else:
        att_loc = [''] * len(att_name)
    for i, attName in enumerate(att_name):
        try:
            attLoc = att_loc[i]
            attVal = att_val[i]
        except:
            sys.exit("\nERROR There must be as many Attribute Location" \
                   + "and Values as Attribute Names (comma separated even if null)")
        # Attribute is added at xpath location unless Attribute Location col
        # is filled
        if attLoc:
            try:
                xpath_list = xpath.split("/")[:]
                loc = xpath_list.index(attLoc)
                xpath_list = xpath_list[:loc+1]
                xpath = "/".join(xpath_list)
            except ValueError:
                print "WARNING : %s not found in XPATH" % attLoc
                print "Attribute %s added at xpath location %s" % (attName, xpath)
        element = tree.xpath(xpath, namespaces=namespaces)[0] 
        # is there a prefix
        tagsplit = str(attName).split(':')
        if len(tagsplit) == 2:
            prefix, att_name = tagsplit
            element.attrib["{" + namespaces[prefix] + "}" + att_name] = attVal
        else:
            att_name = str(attName)
            element.attrib[att_name] = attVal

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
def addNewElementAndValue(tree, tag_list, value, parent_xpath, attribute='No', isAttVal=True):
    for tag in tag_list:
        parent = tree.xpath(parent_xpath, namespaces=namespaces)[0]
        prefix, tag_name = str(tag).split(':')
        new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name)
        parent.insert(0, new_element)
        parent_xpath += "/" + tag
        if tag == tag_list[-1] and value:
           if isAttVal:
           # add the same value for attribute and tag value (like codelist)
               new_element.text = value.strip()
           if attribute != 'No':
                new_element.attrib[attribute] = value
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
def addDateType(tree, xpath, type, code_list):
    # add the date type : creation, publication or revision
    # written after "Date:"
    value = type.split(':')[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:dateType', 'gmd:CI_DateTypeCode']
    xpath = "/".join(xpath_list)
    addMetadataElement(tree, xpath, value)
    addMetadataElement(tree, xpath, value, 'codeListValue')
    addMetadataElement(tree, xpath, code_list, 'codeList')

# Add KeywordType value and the associated codelist 
def addKeywordType(tree, xpath, type, code_list):
    # add the Keyword type written after Keyword
    value = type.split(':')[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:type', 'gmd:MD_KeywordTypeCode']
    xpath = "/".join(xpath_list)
    addMetadataElement(tree, xpath, value)
    addMetadataElement(tree, xpath, value, 'codeListValue')
    addMetadataElement(tree, xpath, code_list, 'codeList')

# Add resource format name
# and associated link, version and mimeType
# Multiple format can be specified separated by ;
# name, version, specification, mimeType are comma separated
def addResourceFormat(tree, xpath, value, urn):
    # initialisation of lists containing format information
    name_list = [] ; version_list = [] ; spec_list = [] ; mime_list = []
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
            mime = val[3].strip()
        except IndexError:
            print val
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
        name_list.append(name)
        version_list.append(version)
        spec_list.append(specification)
        mime_list.append(mime)
        return name_list, version_list, spec_list, mime_list
