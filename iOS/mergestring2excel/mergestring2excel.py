#!/usr/bin/python

#usage: ./xl2.py excel_file language Localizable.strings

import xml.etree.ElementTree as ET
import sys
import xlrd
import xlwt
import string
import collections
import re

sys.path.append('./util')
import excel2listutil
import string2listutil
import helperutil
import list2excelutil

reload(sys)
sys.setdefaultencoding('utf-8')

# export string file to excel file
def exportstring2excel(stringlist, excellist, newfilename):    
    excelValuelist = []
    for k, v in excellist:
        excelValuelist.append(v)
    
    tindex = -1
    for k, v in stringlist:
        tindex = helperutil.kexistlist(v, excelValuelist)
        if tindex >= 0:
            excellist[tindex] = [k, v]
        else:
            excellist.append([k, v])
    
    list2excelutil.list2excel(excellist, newfilename)

def main():
    if len(sys.argv) != 6:
        print 'usage: %s input_excel_file sheetname language Localizable.strings output_excel_file\n' % (sys.argv[0])
        sys.exit()
    excellist = excel2listutil.excel2list(sys.argv[1], sys.argv[2], sys.argv[3])
    stringlist = string2listutil.string2list(sys.argv[4])
    exportstring2excel(stringlist, excellist, sys.argv[5])
    
    print 'success merage string file to excel file'

if __name__ == '__main__':
    main()