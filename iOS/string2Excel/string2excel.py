#!/usr/bin/python

import xml.etree.ElementTree as ET
import sys
import xlrd
import xlwt
import sys
import string
import collections

reload(sys)
sys.setdefaultencoding('utf-8')

def xml2list(file_name):
    f1 = open(file_name)
    fp3 = open("Duplicate.strings", "w")
    f1_list = []
    tmp = []
    duplicatestring = 0
    
    for line in f1:
        lst = line.split('=')
        if len(lst) == 0:
            continue
        key = lst[0].strip('" \n')
        if len(key) == 0:
            continue
        if key.startswith('//'):
            continue
        
        if len(lst) == 2:
            val = lst[1].strip('"; \n')
            tmp = [key, val]
            if tmp not in f1_list:
                f1_list.append(tmp)
            else:
                duplicatestring = 1
                s = '\"%s\" = \"%s\";\n' % (key, val)
                fp3.write(s.decode("utf-8"))
                print 'Duplicate in strings k=%s, v=%s' % (key, val)
    
    if duplicatestring == 1:
        print 'error !!!'
        sys.exit()
    
    return f1_list

def list2excel(lst, filename):
#    wb = xlwt.Workbook()
    wb = xlwt.Workbook(encoding = 'utf-8')
    ws = wb.add_sheet('Sheet Export')
#    for i in range(0, len(f_dict)):
    i = 0
    for k, v in lst:
        ws.row(i).write(0, k)
        ws.row(i).write(1, v)
        i += 1
    wb.save(filename + '.xls')

if __name__ == '__main__':
    if len(sys.argv) != 3:
        print 'usage: %s Localizable.strings excel_file' % (sys.argv[0])
        sys.exit()
    list = xml2list(sys.argv[1])
    list2excel(list, sys.argv[2])
    print 'success export strings file to excel'


