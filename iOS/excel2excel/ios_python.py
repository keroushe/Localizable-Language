#!/usr/bin/python
#usage: ./xl2.py excel_file language Localizable.strings
import sys
import string
import pdb
import collections
import xlwt
import xlrd
from xlutils.copy import copy

reload(sys)
sys.setdefaultencoding('utf-8')

def sheet2list(sh, lans, file, shname):
    """get sheet all data
    :param sh:sheet
    :return:list [{"key":"APP_Cancel", "val":rowdict}]
    rowdict -> {"zh":val, "en":val}
    """
    list = []
    nrows = sh.nrows
    ncols = sh.ncols
    """get first lan rows"""
    for col in range(0, ncols):
        s = '%s' % (sh.cell_value(0, col))
        lan = s.strip().lower()
        lans.append(lan)

    # keys = []
    for row in range(1, nrows):
        key = '%s' % (sh.cell_value(row, 0))
        if len(key) == 0:
            continue
        # if key not in keys:
        #     keys.append(key)
        # else:
        #     print 'Duplicate key %s in file:%s, shname:%s' % (key, file, shname)
        #     sys.exit()

        rowdict = {}
        for col in range(0, ncols):
            s = '%s' % (sh.cell_value(row, col))
            rowdict[lans[col]] = s.strip()
        rowkv = {}
        rowkv['key'] = key
        rowkv['val'] = rowdict
        list.append(rowkv)

    return list

def checksheetlans(lanlist, lans, file):
    """check all sheet lans"""
    listindex = 0
    for landict in lanlist:
        shname = landict["shname"]
        shlans = landict["shlans"]
        if listindex == 0:
            preshlans = shlans
        else:
            if len(preshlans) != len(shlans):
                print 'lans len not same, shname = %s\n, shlanslen = %d, shlans = %s\n, preshlanslen = %d, preshlans = %s\n in file:%s' % (shname, len(shlans), shlans, len(preshlans), preshlans, file)
                sys.exit()

            for i in range(0, len(preshlans)):
                if cmp(shlans[i], preshlans[i]) != 0:
                    print 'lans len not same, shname = %s\n, shlanslen = %d, shlans = %s\n, preshlanslen = %d, preshlans = %s\n in file:%s' % (shname, len(shlans), shlans, len(preshlans), preshlans, file)
                    sys.exit()
        listindex += 1

    if len(lanlist) >= 1:
        lans.extend(lanlist[0]["shlans"])


def excel2list(rb, lans, file):
    """
    convert excel data to list
    :param rb:rb
    :return:[shlist, shlist, shlist]
    """
    list = []
    lanlist = []

    nsheets = len(range(rb.nsheets))
    for page in range(0, nsheets):
        sh = rb.sheet_by_index(page)
        shname = rb.sheet_names()[page]
        shlans = []
        shlist = sheet2list(sh, shlans, file, shname)
        list.extend(shlist)
        landict = {}
        landict["shname"] = shname
        landict["shlans"] = shlans
        lanlist.append(landict)

    checksheetlans(lanlist, lans, file)

    return list

def excel2pagelist(rb, lans, file):
    """
    convert excel data to list by page sheet
    :param rb:rb
    :return:[{"shname":name, "shval": shlist}, {"shname":name, "shval": shlist}]
    """
    list = []
    lanlist = []

    nsheets = len(range(rb.nsheets))
    for page in range(0, nsheets):
        sh = rb.sheet_by_index(page)
        shname = rb.sheet_names()[page]
        shlans = []
        shlist = sheet2list(sh, shlans, file, shname)
        shdict = {}
        shdict['shname'] = shname
        shdict['shval'] = shlist
        list.append(shdict)
        """lans"""
        landict = {}
        landict["shname"] = shname
        landict["shlans"] = shlans
        lanlist.append(landict)

    checksheetlans(lanlist, lans, file)

    return list

def findlanval(key, lan, fromlist):
    val = ''
    for rowkv in fromlist:
        k = rowkv['key']
        rowdict = rowkv['val']
        if cmp(k, key) == 0:
            if rowdict.get(lan):
                val = rowdict.get(lan)
                return val

    return val

def isneedwriterow(key):
    """
    judge current row is need to replace and write
    :param key:lan key
    :return:isneed
    """
    isneed = 1
    if cmp(key.lower(), 'fields') == 0:
        # print '%s is not lan, skip' % (k)
        isneed = 0
    if cmp(key.lower(), 'androidisused') == 0 or \
            cmp(key.lower(), 'iosisused') == 0 or \
            cmp(key.lower(), 'androidiosnoused') == 0 or \
            cmp(key.lower(), 'modestatus') == 0:
        # print '%s is not lan, skip' % (k)
        isneed = 0
    if cmp(key.lower(), 'maxlength') == 0:
        # print '%s is not lan, skip' % (k)
        isneed = 0

    return isneed


def writepagelist2file(wb, list, lans, file):
    for page in range(0, len(list)):
        shdict = list[page]
        shname = shdict['shname']
        shlist = shdict['shval']
        # print 'shname = %s, page = %d' % (shname, page)
        ws = wb.get_sheet(page)
        if not ws:
            print 'sheet not exist, exception !!!'
            sys.exit()

        """write lans"""
        for i in range(0, len(lans)):
            # print 'write lan : %s' % (lans[i])
            ws.row(0).write(i, lans[i])
        """write row by row"""
        for r in range(0, len(shlist)):
            rowkv = shlist[r]
            key = rowkv['key']
            """write key"""
            ws.row(r+1).write(0, key)
            rowdict = rowkv['val']  # this val is dict, and key is lan, val is multilan
            for cow in range(0, len(lans)):
                if isneedwriterow(lans[cow]) == 0:
                    continue
                ws.row(r+1).write(cow, rowdict[lans[cow]])

    wb.save(file)

def exportexcel2excel(fromfile, tofile):
    fromrblans = []
    torblans = []

    fromrb = xlrd.open_workbook(fromfile)
    torb = xlrd.open_workbook(tofile)
    towb = copy(torb)
    fromlist = excel2list(fromrb, fromrblans, fromfile)
    tolist = excel2pagelist(torb, torblans, tofile)

    """write data to tolist"""
    for shdict in tolist:
        shname = shdict['shname']
        shlist = shdict['shval']
        print 'start current sheet : %s' % (shname)
        for rowkv in shlist:
            key = rowkv['key']
            rowdict = rowkv['val']  #this val is dict, and key is lan, val is multilan
            for k, v in rowdict.iteritems():
                if isneedwriterow(k) == 0:
                    continue

                """find key of lan from fromlist"""
                newval = findlanval(key, k, fromlist)
                if len(newval) != 0:
                    rowdict[k] = newval
                    # print '%s value for %s is found, newval : %s' % (k, key, newval)
                # else:
                #     print '%s value for %s is not found' % (k, key)

    writepagelist2file(towb, tolist, torblans, tofile)

def main():
    if len(sys.argv) != 3:
        print 'usage: %s from_excel_file to_excel_file' % (sys.argv[0])
        sys.exit()
    exportexcel2excel(sys.argv[1], sys.argv[2])
    print 'success import excel file to strings'

if __name__ == '__main__':
    main()