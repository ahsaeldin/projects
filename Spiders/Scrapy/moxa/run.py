#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'moxa_spider2' #moxa_spider

if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o Moxa_Raw_Outs.csv".format(name).split()
    cmdline.execute(command)