#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'yp_us_scraper'
if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o YP_USA_Outs.csv".format(name).split()
    cmdline.execute(command)