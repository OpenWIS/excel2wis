#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Set library version (A.B.C)
# from excel file version (A.B)
# and python script version (C)

import xlrd
import argparse
import re
import os


# Read excel version
workbook = xlrd.open_workbook('excel2wisxml/templates/Metadata-guide-record.xls')
md_gene = workbook.sheet_by_name('MD generic')
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
for row in range(md_gene_row_start, md_gene.nrows):
    tag = unicode(md_gene.cell_value(row, md_gene_tag_col)).strip()
    value = unicode(md_gene.cell_value(row, md_gene_value_col)).strip()
    if tag.startswith('ExcelVersion'):
        excelVersion = value
        break
del workbook

# Read script version
with open("excel2wisxml/excel2wisxml.py") as f:
    script_lines = f.readlines()
    for line in script_lines:
        if "SCRIPT_VERSION" in line:
            scriptVersion = re.split("\"", line)[1]
            break

# Create application version
app_version = excelVersion + "." + scriptVersion
print "Version %s" % app_version

with open("setup.py", 'r') as f:
    setup_lines = f.readlines()
    for i, line in enumerate(setup_lines):
        if "version" in line:
            setup_lines[i] = "    version='%s',\n" % app_version
            break

with open("setup.py", 'w') as f:
    for line in setup_lines:
        f.write(line)

commit_cmd = 'git add setup.py ; git commit setup.py -m "Version %s in setup.py"' % app_version
os.system(commit_cmd)
