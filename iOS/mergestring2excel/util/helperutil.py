#!/usr/bin/python

import sys
import string

def kexistlist(key, list):
    for i in range(len(list)):
        if list[i] == key:
            return i
    return -1

def printlist(list):
    for v in list:
        print 'v:%s' % (v)

def print_arr_in_list(list):
    for k, v in list:
        print 'k:%s,v:%s' % (k, v)