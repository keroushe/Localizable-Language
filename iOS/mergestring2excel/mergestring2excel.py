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

reload(sys)
sys.setdefaultencoding('utf-8')

# export string file to excel file
def exportstring2excel(stringlist, excellist, newfilename):    
    tmp = []
    stringkeylist = []
    tindex = -1
    for k, v in stringlist:
        stringkeylist.append(k)
    
    for k, v in excellist:
        tindex = kexistlist(k, stringkeylist)
        if tindex >= 0:
            excellist[tindex] = [k,stringlist[tindex]]
        else:
            
            xmllist.append(tmp)
    
    f = open(file_name, 'w')
    for k, v in xmllist:
        s = "\"%s\"=\"%s\";\n" % (k, v)
        f.write(s)
    f.flush()
    f.close()

def main():
    if len(sys.argv) != 5:
        print 'usage: %s excel_file sheetname language Localizable.strings\n' % (sys.argv[0])
        sys.exit()
    excellist = excel2listutil.excel2list(sys.argv[1], sys.argv[2], sys.argv[3])
    stringlist = string2listutil.string2list(sys.argv[4])

    
    print 'success import excel file to strings'

if __name__ == '__main__':
    main()