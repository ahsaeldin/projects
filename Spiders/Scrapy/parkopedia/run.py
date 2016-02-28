#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'parkopedia_scraper_by_urls'
if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o parkopedia_outs.csv".format(name).split()
    cmdline.execute(command)