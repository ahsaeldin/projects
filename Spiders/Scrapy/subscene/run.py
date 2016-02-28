#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'SubsceneCrawler'
if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o subscenev3_outs.csv".format(name).split()
    cmdline.execute(command)