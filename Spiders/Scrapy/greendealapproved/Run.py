#!/usr/bin/python

import sys

from scrapy import cmdline

if __name__ == '__main__':
    cmdline.execute("scrapy crawl greendealapproved -o greendealapproved-data.csv -t csv".split())