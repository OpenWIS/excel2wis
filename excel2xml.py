#!/usr/bin/env python
# -*- coding: utf-8 -*-

import copy
import sys
import xlrd
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

def addMultiValue(tree, xpath, multivalue):
    xpath_list = xpath.split("/")[1:]
    xpath = ""
    # Rebuild of xpath to add missing tag
    for i, tag in enumerate(xpath_list):
        previous_xpath = xpath
        xpath += "/" + tag
        try:
            element = tree.xpath(xpath, namespaces=namespaces)
        except etree.XPathEvalError:
            if tag.endswith('[]'):
                break
        print "addMutlivalue, après le try", element
        if len(element) == 0:
            # element under which the tag will be added
            parent = tree.xpath(previous_xpath, namespaces=namespaces)[0]
            prefix, tag_name = str(tag).split(':')
            sub_element = etree.SubElement(parent, "{" + namespaces[prefix] + "}" + tag_name)
    print i, tag
    # MultiValue but no [] found in xpath
    if not tag.endswith('[]'):
        raise ValueError
    multi_tag_list = xpath_list[i:]
    multi_tag_list[0] = tag[:-2]
    for val in reversed(multivalue.split(',')):
        parent_xpath = previous_xpath
        for tag in multi_tag_list:
            parent = tree.xpath(parent_xpath, namespaces=namespaces)[0]
            prefix, tag_name = str(tag).split(':')
            new_element = parent.makeelement("{" + namespaces[prefix] + "}" + tag_name)
            parent.insert(0, new_element)
            parent_xpath += "/" + tag
            if tag == multi_tag_list[-1]:
                new_element.text = val.strip()

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
            previous_xpath = xpath
            xpath += "/" + tag
            element = tree.xpath(xpath, namespaces=namespaces)
            # missing tag identified
            if len(element) == 0:
                # element under which the tag will be added
                parent = tree.xpath(previous_xpath, namespaces=namespaces)[0]
                prefix, tag_name = str(tag).split(':')
                sub_element = etree.SubElement(parent, "{" + namespaces[prefix] + "}" + tag_name)
        el = sub_element
    # Insert tag or attribute value
    if attribute == 'No':
        el.text = value
    else:
        el.attrib[attribute] = value

def concateValue(tree, value):
    # Unique Identifier (value for 1.3)
    urn = unicode(md_gene.cell_value(5, 2)).strip() + value
    # Location for online access
    nrow = 4
    value = unicode(md_gene.cell_value(nrow, 2)).strip() + urn
    xpath = unicode(md_gene.cell_value(nrow, 3)).strip()
    addMetadataElement(tree, xpath, value)
    # Permanent link
    nrow = 3
    value = unicode(md_gene.cell_value(nrow, 2)).strip() + urn
    xpath = unicode(md_gene.cell_value(nrow, 3)).strip()
    addMetadataElement(tree, xpath, value)
    # Two linked tags are mandatory, cf. template (paragraph4)
    return urn

def par1022(tree, xpath, name):
    # dateType change (cf. template paragraph10) according to which column is filled among (5.2, 5.3, 5.4)
    value = name.split()[-1]
    xpath_list = xpath.split("/")[:-2] + ['gmd:dateType', 'gmd:CI_DateTypeCode']
    xpath = "/".join(xpath_list)
    addMetadataElement(tree, xpath, value)
    addMetadataElement(tree, xpath, value, 'codeListValue')

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

# Get sheets
md_fields = workbook.sheet_by_name('MD Fields')
help = workbook.sheet_by_name('Help')

###
# Track FATAL ERRORS
###
field_id_list = md_fields.row(3)
field_mandatory_list = md_fields.row(4)
help_id_list = help.col(1)
try:
    for i, id in enumerate(field_id_list):
        field_id = unicode(id.value).strip()
        help_id = unicode(help_id_list[i+2].value).strip()
        # Check if MD Fields ID and Help ID match
        if field_id != help_id:
            raise Exception(
                    "ERROR : Paragraphs number in MD Fields sheet : "
                    + "%s doesn't match paragraphs" % field_id
                    + "number in Help sheet : %s" % help_id)
        # Check (non INSPIRE) mandatory fields
        # TODO INSPIRE : check mandatory fields for INSPIRE ?
        mandatory = unicode(field_mandatory_list[i].value).strip()
        if mandatory == 'Mandatory' and i < 29:
            for row in range(6, md_fields.nrows):
                # mandatory field is not empty
                if not md_fields.cell_value(row, i): 
                    raise Exception('ERROR - MD Fields sheet - ' +
                                    'Mandatory field %s' % field_id +
                                    ' is empty on row %s' % str(row+1))
                # xpath linked to mandatory field is not empty
                if not help.cell_value(i+2, 8):
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

###
# Add generic metadata (MD generic sheet)
###
# Get MD generic sheet
md_gene = workbook.sheet_by_name('MD generic')
# Lists for WARN messages
empty_xpath_gene = []
error_gene = []
for row in range(2, md_gene.nrows):
    value = unicode(md_gene.cell_value(row, 2)).strip()
    xpath = unicode(md_gene.cell_value(row, 3)).strip()
    code_list = unicode(md_gene.cell_value(row, 4)).strip()
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
        error_gene.append(row)
        continue

# Write WARN messages for MD generic
if empty_xpath_gene or error_gene:
    print "\n--- WARN -------- MD generic ---"
    if empty_xpath_gene:
        print "elements on row(s) %s have no XPATH" % ", ".join([str(x) for x in empty_xpath_gene])
    if error_gene:
        print "elements on row(s) %s cannot be created, please check their xpath expression" % ", ".join([str(x) for x in error_gene])
    print "--------------------------------\n"

###
# Add specific metadata (MD Fields and Help sheets)
###
# Iteration on MD Fields rows (one row = one metadata)
for row in range(6, md_fields.nrows):
    tree = copy.deepcopy(common_tree)
    # Lists for WARN messages
    empty_xpath = []
    error = []
    for col in range(1, md_fields.ncols):
        id = unicode(field_id_list[col].value).strip()  # element ID
        # TODO : translation
        if id.endswith('b'):
            print "> translation %s ignored" % id
            continue
        # An optional empty field is not added
        mandatory = unicode(field_mandatory_list[col].value).strip()
        if mandatory == 'Optional':
            if not md_fields.cell_value(row, col): 
                print "> empty optional field %s ignored" % id 
                continue    
        field_value = unicode(md_fields.cell_value(row, col)).strip()
        xpath = unicode(help.cell_value(col+2, 8)).strip()
        attribute = unicode(help.cell_value(col+2, 4)).strip()
        multivalue = unicode(help.cell_value(col+2, 6)).strip()
        code_list = unicode(help.cell_value(col+2, 7)).strip()
        # empty Xpath
        if xpath == '':
            empty_xpath.append(id)
            continue
        try:
            #print "ID", id
            if id == '1.3':
                # add tag values which are concatenation of MD generic and MD Fields elements
                field_value = concateValue(tree, field_value)
            if multivalue == 'Yes':
                addMultiValue(tree, xpath, field_value)
            else:
                addMetadataElement(tree, xpath, field_value, attribute)
            if id in ['5.2', '5.3', '5.4']:
                # add creation, publication or revision in dateType (paragraph 10.2.2)
                name = unicode(md_fields.cell_value(5, col)).strip()
                par1022(tree, xpath, name)
            if code_list:
                addMetadataElement(tree, xpath, field_value, 'codeListValue')
                addMetadataElement(tree, xpath, code_list, 'codeList')
        except Exception as e:
            print id, type(e), e
        except ValueError:
            error.append(id)
            continue
        except etree.XPathEvalError:  # [] in xpath
            error.append(id)

    # Write an xml file for each metadata (row in MD Fields)
    metadata_row = row + 1
    string_xml = etree.tostring(tree, pretty_print=True)
    filename = "metadata_row" + str(metadata_row) + ".xml"
    with open(filename, "wb") as fo:
        fo.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        fo.write(string_xml)

    print "\n##### File %s has been generated\n" % filename

    # Write WARN messages for MD Fields - Help for each row
    if empty_xpath or error:
        print "--- WARN -------- Fields row", metadata_row, "- Help ---"
        if empty_xpath:
            # not empty (optional) elements in MD Fields with no xpath linked in Help
            print "elements %s have no XPATH" % ", ".join(empty_xpath)
        if error:
            print "elements %s cannot be created, please check their xpath expression" % ", ".join(error)
    print "-----------------------------------------\n"
