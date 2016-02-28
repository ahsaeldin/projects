#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'aido'
output_file_name = 'aido-outputs.csv'

if __name__ == '__main__':
    command = "scrapy crawl {0} -t csv -o Aido-Outputs.csv.csv".format(name).split()
    cmdline.execute(command)