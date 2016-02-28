#!/usr/bin/env python

import os
from random import choice

def SaveLstIntoFile(prodURLs, fileName):
    iFile = open(fileName, "ab")
    for prodURL in prodURLs:
        iFile.write(prodURL + '\r\n')
    iFile.close()

def loadProxies(fileName):
    return loadFileIntoLIst(fileName)

def loadFileIntoLIst(fileName):
    with open(fileName) as f:
        _proxies = f.read().splitlines()
    return _proxies

def PickRandomProxy(proxies):
    pickOne = choice(proxies)
    return pickOne

def removeDataFile(dataFileName):
    try:
        os.remove(dataFileName)
    except OSError:
        pass