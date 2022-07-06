#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'dillards'
if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o outs.csv".format(name).split()
    cmdline.execute(command)

