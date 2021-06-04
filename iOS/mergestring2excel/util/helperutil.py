#!/usr/bin/python

import sys
import string

def kexistlist(key, list):
        for i in range(len(list)):
            if list[i] == key:
                return i
        return -1