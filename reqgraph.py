#!/usr/bin/env python3

import json
import sys

class Req(object):

    def __init__(self, level, num, dsc):

        self.__num     = num
        self.__dsc     = dsc
        self.__level   = level
        self.__refs    = []
        self.__refbys  = []

    def addRefines(self, num, dsc):
        self.__refs.append((num, dsc))

    def addRefinedBy(self, num):
        self.__refbys.append(num)

    def level(self):
        return self.__level

    def num(self):
        return self.__num

    def dsc(self):
        return self.__dsc

    def refines(self):
        return self.__refs

    def refinedBy(self):
        return self.__refbys


def link(reqs):

    # Link the parents to the children 
    for r in reqs:

        for re, dsc in reqs[r].refines():

            if reqs[re].level() == reqs[r].level()-1:
                reqs[re].addRefinedBy(reqs[r])
            else:
                print('WARNING: Requirement %s at level %d cannot refine %s at level %d.' %(reqs[r].num(), reqs[r].level(), reqs[re].num(), reqs[re].level()))

    return reqs

def load(files):    

    reqs = {}

    for filename in files:
        with open(filename, 'r') as f:
            j = json.load(f)
            for r in j['reqs']:
                reqs[r['num']] = Req(j['level'], r['num'], r['dsc'])

                try:

                    # Catches the Index error (no 'ref') element
                    # Also catches the empty list error.
                    r['ref'][0]
                    for re in r['ref']:
                        reqs[r['num']].addRefines(re['num'], re['dsc'])


                except (KeyError, IndexError):
                    if reqs[r['num']].level() != 0:
                        print('WARNING: Requirement %s is an orphan. Add a ref element.' %(r['num']))

    return reqs

def build(files):

    return link(load(files))

