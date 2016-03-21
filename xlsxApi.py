#!/usr/bin/python
# -*- coding: UTF-8 -*-

import string
import sys
import os
import re
import xlsxwriter
import datetime
import cgi
import cgitb
cgitb.enable()

print "Content-Type: text/plain;charset=utf-8"
print

def strip_non_ascii(string):
	stripped = (c for c in string if 0 < ord(c) < 127)
	return ''.join(stripped)

form = cgi.FieldStorage()
xlsPath = form.getvalue("folder")
xlsFile = form.getvalue("file")

xlsBasepath = "/shares/api_xlsx/"
xlsInput = xlsBasepath+xlsPath+"/"+xlsFile
xlsxTemp = xlsFile.replace(".xls",".xlsx")
xlsxName = xlsBasepath+xlsPath+"/"+xlsxTemp

workbook = xlsxwriter.Workbook(xlsxName, {'strings_to_numbers': True})
worksheet = workbook.add_worksheet()
workbook.set_properties({
    'title':    'Micromedia report',
    'author':   'Micromedia B.V.',
    'company':  'Micromedia B.V.'})

# Define date formatting and create regular expression
dateFormat = workbook.add_format({'num_format': 'dd-mm-yyyy'})
dateFormEU = "%d-%m-%Y"
dateFormAM = "%Y-%m-%d"
regexpDateEU = re.compile(r'^(0[1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d')
regexpDateEU2 = re.compile(r'^([1-9]|[12][0-9]|3[01])[- /.]([1-9]|1[012])[- /.](19|20)\d\d')
regexpDateEU3 = re.compile(r'^(0[1-9]|[12][0-9]|3[01])[- /.]([1-9]|1[012])[- /.](19|20)\d\d')
regexpDateEU4 = re.compile(r'^([1-9]|[12][0-9]|3[01])[- /.](0[1-9]|1[012])[- /.](19|20)\d\d')
regexpDateAM = re.compile(r'^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])')
regexpDateAM2 = re.compile(r'^(19|20)\d\d[- /.]([1-9]|1[012])[- /.]([1-9]|[12][0-9]|3[01])')
regexpDateAM3 = re.compile(r'^(19|20)\d\d[- /.](0[1-9]|1[012])[- /.]([1-9]|[12][0-9]|3[01])')
regexpDateAM4 = re.compile(r'^(19|20)\d\d[- /.]([1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])')

# Set rows and columns to 0 to start from cell A1
rows = 0
cols = 0
file = open(xlsInput)
# Loop through input file line by line
with file as inputValues:
    for line in inputValues:
		# Split the lines by tab
        values = line.split("\t")
        for val in values:
			# Strip all non-ascii values
			val = strip_non_ascii(val)
			val = val.replace(',','.')
			val = val.replace('"','')
			# Check if value is date and format accordingly
			if regexpDateEU.search(val) is not None or regexpDateEU2.search(val) is not None or regexpDateEU3.search(val) is not None or regexpDateEU4.search(val) is not None:
				val = val.replace('/','-')
				val = val.replace("\r",'')
				val = val.replace("\n",'')
				val = val.split(" ")[0]
				val = datetime.datetime.strptime(val, dateFormEU)
				# Write to the worksheet
				worksheet.write(rows, cols, val, dateFormat)
			elif regexpDateAM.search(val) is not None or regexpDateAM2.search(val) is not None or regexpDateAM3.search(val) is not None or regexpDateAM4.search(val) is not None:
				val = val.replace('/','-')
				val = val.replace("\r",'')
				val = val.replace("\n",'')
				val = val.split(" ")[0]
				val = datetime.datetime.strptime(val, dateFormAM)
				# Write to the worksheet
				worksheet.write(rows, cols, val, dateFormat)
			else:
				# Write to the worksheet
				worksheet.write(rows, cols, val)
			cols += 1
        cols = 0
        rows += 1

workbook.close()
os.unlink(xlsInput)