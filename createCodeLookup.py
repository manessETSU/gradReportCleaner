import xlrd
import xlwt
import sys
import re


#CIPC

print "Opening Workbook" + str(sys.argv[1]) + " and sheet " + str(sys.argv[2])

workbook = xlrd.open_workbook(sys.argv[1])
worksheet = workbook.sheet_by_name(str(sys.argv[2]) )

num_cells = worksheet.ncols - 1
curr_cell = 0

CIPCColumnIndex = -1

headerDictionary = {}

while curr_cell < num_cells:
  cell_value = worksheet.cell_value(0, curr_cell)
  headerDictionary[cell_value] = curr_cell
  curr_cell+=1

CIPCColumnIndex = headerDictionary["CIPC"]

curr_cell = 1
num_cells = worksheet.nrows - 1

while curr_cell < num_cells:

  line = str(worksheet.cell_value(curr_cell, CIPCColumnIndex))
  line = line.split(".")[0]
  if line != " ":
    row = worksheet.row(curr_cell)
    print row[0] + " " + row[1]

  curr_cell+=1
