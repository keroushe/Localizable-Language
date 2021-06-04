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

# parse strings file
def string2list(file_name):
    f1 = open(file_name)
    fp3 = open("Duplicate_strings.strings", "w")
    f1_list = []
    keys = []
    tmp = []
    duplicatestring = 0
    pattern = re.compile(r'"(\s|\S){0,}?[^\\]";{0,1}')

    for line in f1:
        lst = []
        for match in re.finditer(pattern, line):
            s = match.start()
            e = match.end()
            lst.append(line[s:e])

        # lst = line.split('=')
        if len(lst) <= 1:
            continue
        elif len(lst) != 2:
            print '%s Format error, please import manually !!!' % (lst[0])
            continue
        
        key = lst[0].strip()
        val = lst[1].strip()

        key = key.strip('"')
        if val.endswith('\"";'):
            val = val[1:-2]
        else:
            val = val.strip('";')
        
        if len(key) == 0:
            continue
        if key.startswith('//'):
            continue
        
        if key not in keys:
            keys.append(key)
        else:
            duplicatestring = 1
            s = '\"%s\" = \"%s\";\n' % (key, val)
            fp3.write(s.decode("utf-8"))
            print 'Duplicate in strings key=%s, value=%s' % (key, val)

        # tmp = [key, val]
        tmp = [val, key]
        if tmp not in f1_list:
            f1_list.append(tmp)

    if duplicatestring == 1:
        print 'error !!!'
        sys.exit()

return f1_list