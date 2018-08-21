#!/usr/bin/python

import xlrd
import xlwt
from xlutils.copy import copy
import os

filename='C:\\Users\\Administrator\\Desktop\\hgvs_added.xlsx'
data = xlrd.open_workbook(filename)
table = data.sheets()[0]
data2 = copy(data)
sheet = data2.get_sheet(0)
D = {"A":"T","G":"C","C":"G","T":"A"}
arr = []
for i in range(table.nrows):
    if table.cell(i,9).value == "-":
		arr.append(table.row_value(i))
	    print(arr)

data2.save("C:\\Users\\Administrator\\Desktop\\new.xls")
