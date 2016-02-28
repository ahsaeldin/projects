#!/usr/bin/python

import sys
import time
from scrapy import cmdline

if __name__ == '__main__':
    command = 'scrapy crawl resortdata -o data.csv -t csv'

    cmdline.execute(command.split())
