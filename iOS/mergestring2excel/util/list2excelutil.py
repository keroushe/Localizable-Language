#!/usr/bin/python

import xml.etree.ElementTree as ET
import sys
import xlrd
import xlwt
import string
import collections
import re

reload(sys)
sys.setdefaultencoding('utf-8')

def list2excel(lst, filename):
    wb = xlwt.Workbook(encoding = 'utf-8')
    ws = wb.add_sheet('Sheet Export')
    i = 0
    for k, v in lst:
        ws.row(i).write(0, k)
        ws.row(i).write(1, v)
        i += 1
    wb.save(filename + '.xls')