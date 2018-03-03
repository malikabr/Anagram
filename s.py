__author__ = 'ARAD'
import os
import xlrd,sys
import openpyxl
from xlrd import open_workbook
from openpyxl import load_workbook


#!/usr/bin/env python
# -*- coding: utf-8 -*-


def perms(s):
    if len(s) == 1:
        return [s]
    result = []
    for i, v in enumerate(s):
        result += [v+p for p in perms(s[:i]+s[i+1:])]
    return result

test = input("Enter Your Word Please!")
li = perms(test)
print(li)


from xlrd import open_workbook
book = open_workbook('wordlist.xlsx')
for sheet in book.sheets():
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            for i in range(0, len(li)):
                if cell.value == li[i]:
                    print(cell.value, colidx, rowidx)


#this is a complete program