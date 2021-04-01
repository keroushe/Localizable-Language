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

def get_kv(localizable_file_name):
    f1 = open(localizable_file_name)
    f1_dict = {}
    fp3 = open("Duplicate_strings.strings", "w")
    duplicatestring = 0
    for line in f1:
        lst = line.split('=')
        # pdb.set_trace()
        if len(lst) == 0:
            continue
        key = lst[0].strip('" \n')
        if len(key) == 0:
            continue
        if key.startswith('//'):
            continue
        # pdb.set_trace()
        # print len(lst)
        if len(lst) == 2:
            val = lst[1].strip('"; \n')
            # pdb.set_trace()
            if key not in f1_dict:
                f1_dict[key] = val
            else:
                duplicatestring = 1
                s = '\"%s\" = \"%s\";\n' % (key, val)
                fp3.write(s.decode("utf-8"))
                print 'Duplicate in strings k=%s, v=%s' % (key, val)
    
    if duplicatestring == 1:
        print 'error !!!'
        sys.exit()

    return f1_dict

def xl2dict(excel_file, sheetname, language):
    """get dict from excel, then check is Duplicate in excel
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

    xldict = {}
    fp1 = open("Empty_excel.strings", "w")
    fp2 = open("Duplicate_excel.strings", "w")
    emptystring = 0
    duplicatestring = 0

    for i in range(2, nrows):
        #if str(sh.cell_value(i, col2)).strip() == "":
                    #continue
        #a=sh.cell_value(i, 0)
        #try:
        #    a=int(sh.cell_value(i, 0))
        #except:
        #    pass
        k = str(sh.cell_value(i, 0)).strip()
        v = str(sh.cell_value(i, col)).strip()
        
        if len(k) != 0 and len(v) == 0:
            emptystring = 1
            e = '\"%s\" = \"%s\";\n' % (k, v)
            fp1.write(e.decode("utf-8"))
            print 'Empty in excel \"%s\" = \"%s\"' % (k, v)
        if len(k)  == 0 and len(v) == 0:
            continue
        if k not in xldict:
            xldict[k] = v
        else:
            duplicatestring = 1
            d = '\"%s\" = \"%s\";\n' % (k, v)
            fp2.write(d.decode("utf-8"))
            print 'Duplicate in excel k=%s, v=%s' % (k, v)
    
#    if emptystring == 1 or duplicatestring == 1:
#        print 'error happend'
#        sys.exit()

    if duplicatestring == 1:
        print 'error !!!'
        sys.exit()

    return xldict

def sub(xldict, xmldict, localizable_file_name):
    for k, v in xldict.iteritems():
        if k in xmldict:
            r = v.replace('\\','')
            xmldict[k] = r.replace('"','\\"')
        else:
            r = v.replace('\\','')
            xmldict[k] = r.replace('"','\\"')
    
    f = open(localizable_file_name, 'w')
    d = collections.OrderedDict(sorted(xmldict.items(), key = lambda t: t[0].lower()))
    for k, v in d.iteritems():
        s = "\"%s\" = \"%s\";\n" % (k, v)
        f.write(s)
    f.flush()
    f.close()

def difference(xld, xmld):
    """compare difference between
    excel file and .strings file
    """
    def dcompare(d1, d2):
        tmp = {}
        for k, v in d1.iteritems():
            if k not in d2:
                tmp[k] = v
        return tmp

    def dict2txt(d, filename):
        f = open(filename, "a")
        for k, v in d.iteritems():
            s = "\"%s\" = \"%s\";\n" % (k, v)
            f.write(s.decode("utf-8"))

    filename = "defference.strings"
    with open(filename, "w") as f:
        f.write("\n\n**********Following keys in excel file not in .strings file*********\n\n")
    dict2txt(dcompare(xld, xmld), filename)

    with open(filename, "a") as f:
        f.write("\n\n**********Following keys in .strings file not in excel file*********\n\n")
    dict2txt(dcompare(xmld, xld), filename)

    with open(filename, "a") as f:
        f.write("\n\n**********compare end*********\n\n")

def main():
    if len(sys.argv) != 5:
        print 'usage: %s excel_file sheetname language Localizable.strings\n' % (sys.argv[0])
        sys.exit()
    xld = xl2dict(sys.argv[1], sys.argv[2], sys.argv[3])
    xmld = get_kv(sys.argv[4])
    difference(xld, xmld, sys.argv[4])
    sub(xld, xmld)
    print 'success import excel file to strings'

if __name__ == '__main__':
    main()
