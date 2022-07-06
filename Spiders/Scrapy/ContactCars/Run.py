#!/usr/bin/python

import sys
from scrapy import cmdline

if __name__ == '__main__':
    arg = ''
    args_no = len(sys.argv)
    if args_no > 1:
       arg = sys.argv[1]

    if arg == 'contactcars' or arg == "":
       cmdline.execute("scrapy crawl contactcars -o contactcars.csv -t csv".split())
