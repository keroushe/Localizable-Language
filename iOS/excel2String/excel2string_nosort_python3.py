#!/usr/bin/python
#usage: ./xl2.py excel_file language Localizable.strings
import sys
import xlrd
import sys
import string
import pdb
import collections

reload(sys)
sys.setdefaultencoding('utf-8')

def xml2list(file_name):
    f1 = open(file_name)
    fp3 = open("Duplicate_strings.strings", "w")
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
                print ('Duplicate in strings k=%s, v=%s' % (key, val))
    
    if duplicatestring == 1:
        print ('error !!!')
        sys.exit()

    return f1_list

def excel2list(file_name):
    """get arr from excel, then check is Duplicate in excel
    """
    bk = xlrd.open_workbook(file_name)
    shxrange = range(bk.nsheets)
    sh = bk.sheet_by_index(0)		#open first sheet
    nrows = sh.nrows
    ncols = sh.ncols
    isfound = 0
    for col in range(0, ncols):
        if sh.cell_value(0, col).strip().lower() == sys.argv[2].lower():
            isfound = 1
            break
    if isfound == 1:
        print ('Found lang \"%s\" in %d col' % (sh.cell_value(0, col), col))
    else:
        print ('Not found lang \"%s\"' % (sys.argv[2]))
        sys.exit()

    xllist = []
    tmp = []
    fp1 = open("Empty_excel.strings", "w")
    fp2 = open("Duplicate_excel.strings", "w")
    emptystring = 0
    duplicatestring = 0

    for i in range(2, nrows):
        k = str(sh.cell_value(i, 0)).strip()
        v = str(sh.cell_value(i, col)).strip()
        tmp = [k, v]
        
        if len(k) != 0 and len(v) == 0:
            emptystring = 1
            e = '\"%s\" = \"%s\";\n' % (k, v)
            fp1.write(e.decode("utf-8"))
            print ('Empty in excel \"%s\" = \"%s\"' % (k, v))
        if len(k) == 0:
            continue
        if len(k) == 0 and len(v) == 0:
            continue
        if tmp not in xllist:
            xllist.append(tmp)
        else:
            duplicatestring = 1
            d = '\"%s\" = \"%s\";\n' % (k, v)
            fp2.write(d.decode("utf-8"))
            print ('Duplicate in excel k=%s, v=%s' % (k, v))
    
#    if emptystring == 1 or duplicatestring == 1:
#        print ('error happend')
#        sys.exit()

    if duplicatestring == 1:
        print ('error !!!')
        sys.exit()

    return xllist

def import2xml(excellist, xmllist, file_name):
    def kexistlist(key, list):
        for i in range(len(list)):
            if list[i] == key:
                return i
        return -1
    
    tmp = []
    xmlkeylist = []
    tindex = -1
    for k, v in xmllist:
        xmlkeylist.append(k)
    
    for k, v in excellist:
#        t = v.replace('\\','')
#        v = t.replace('"','\\"')
        v = v.replace('"','\\"')
        tmp = [k, v]
        tindex = kexistlist(k, xmlkeylist)
        if tindex >= 0:
            xmllist[tindex] = tmp
        else:
            xmllist.append(tmp)
    
    f = open(file_name, 'w')
    for k, v in xmllist:
        s = "\"%s\"=\"%s\";\n" % (k, v)
        f.write(s)
    f.flush()
    f.close()

def difference(excellist, xmllist):
    """compare difference between
    excel file and .strings file
    """
    def dcompare(d1, d2):
        ret = []
        tmp = []
        for k, v in d1:
            tmp = [k, v]
            if tmp not in d2:
                ret.append(tmp)
        return ret

    def list2txt(d, filename):
        f = open(filename, "a")
        for k, v in d:
            s = "\"%s\" = \"%s\";\n" % (k, v)
            f.write(s.decode("utf-8"))

    filename = "defference.strings"
    with open(filename, "w") as f:
        f.write("\n\n**********Following keys in excel file not in .strings file*********\n\n")
    list2txt(dcompare(excellist, xmllist), filename)

    with open(filename, "a") as f:
        f.write("\n\n**********Following keys in .strings file not in excel file*********\n\n")
    list2txt(dcompare(xmllist, excellist), filename)

    with open(filename, "a") as f:
        f.write("\n\n**********compare end*********\n\n")

def main():
    if len(sys.argv) != 4:
        print ('usage: %s excel_file language Localizable.strings\nconflict in dict_conflict.txt' % (sys.argv[0]))
        sys.exit()
    excellist = excel2list(sys.argv[1])
    xmllist = xml2list(sys.argv[3])
    difference(excellist, xmllist)
    import2xml(excellist, xmllist, sys.argv[3])
    print ('success import excel file to strings')

if __name__ == '__main__':
    main()
