#!/usr/bin/env python

import openpyxl, sys

try:
	excelDoc = sys.argv[1]
	excelSheet = sys.argv[2]
	colNum = sys.argv[3]
	rowNum = sys.argv[4]
	txtDoc = sys.argv[5]
except IndexError:
	excelDoc = input('Which Excel document am I working with? (absolute path)')
	excelSheet = input('What is the relevant Excel worksheet called?')
	colNum = input('How many columns does it have? (integer)')
	rowNum = input('How many rows does it have? (integer)')
	txtDoc = input('What do you want the .txt output called? (absolute path)')

wb = openpyxl.load_workbook(excelDoc)
sheet = wb[excelSheet]

currentRow = []
dataSet = []

for r in range(1, int(rowNum) + 1):
	
	for c in range(1, int(colNum) + 1):
		currentCell = sheet.cell(row=r, column=c).value
		currentRow.append(currentCell)		
		if c == int(colNum):
			dataSet.append(currentRow)
			currentRow = []

file = open(txtDoc, 'w')

for row in dataSet:
	file.write('  * - ' + str(row[0]))
	
	for cell in range(1, int(colNum)):
		file.write('\n' + '    - ' + str(row[cell]))
	
	file.write('\n')

file.close()