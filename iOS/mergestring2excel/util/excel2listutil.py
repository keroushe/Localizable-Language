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

# parse excel file
def excel2list(excel_file, sheetname, language):
    """get arr from excel, then check is Duplicate in excel
    """
    bk = xlrd.open_workbook(excel_file)
    sh = bk.sheet_by_name(sheetname)
    if sh == None:
        print 'error !!!'
        sys.exit()

    nrows = sh.nrows
    ncols = sh.ncols
    isfound = 0
    for col in range(0, ncols):
        if sh.cell_value(0, col).strip().lower() == language.lower():
            isfound = 1
            break
    if isfound == 1:
        print 'Found lang \"%s\" in %d col' % (sh.cell_value(0, col), col)
    else:
        print 'Not found lang \"%s\"' % (language)
        sys.exit()

    xllist = []
    tmp = []
    keys = []
    fp1 = open("Empty_excel.strings", "w")
    fp2 = open("Duplicate_excel.strings", "w")
    emptystring = 0
    duplicatestring = 0

    for i in range(1, nrows):
        k = str(sh.cell_value(i, 0))
        v = str(sh.cell_value(i, col))
        tmp = [k, v]
        
        if len(k) != 0 and len(v) == 0:
            emptystring = 1
            e = '\"%s\" = \"%s\";\n' % (k, v)
            fp1.write(e.decode("utf-8"))
            print 'Empty in excel \"%s\" = \"%s\"' % (k, v)
        if len(k) == 0:
            continue
        if len(k) == 0 and len(v) == 0:
            continue
        
        if k not in keys:
            keys.append(k)
        else:
            duplicatestring = 1
            d = '\"%s\" = \"%s\";\n' % (k, v)
            fp2.write(d.decode("utf-8"))
            print 'Duplicate in excel k=%s, v=%s' % (k, v)

        if tmp not in xllist:
            xllist.append(tmp)
    
#    if emptystring == 1 or duplicatestring == 1:
#        print 'error happend'
#        sys.exit()

    if duplicatestring == 1:
        print 'error !!!'
        sys.exit()

    return xllist