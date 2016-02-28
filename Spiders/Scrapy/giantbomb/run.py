#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'games_scraper'
if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o Outputs.csv".format(name).split()
    #command = "scrapy crawl {0}".format(name).split()
    cmdline.execute(command)