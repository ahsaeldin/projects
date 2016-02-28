#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys
from scrapy import cmdline

if __name__ == '__main__':
    arg = ''
    args_no = len(sys.argv)

    cmdline.execute("scrapy crawl rpycc -o rpycc.csv -t csv".split())