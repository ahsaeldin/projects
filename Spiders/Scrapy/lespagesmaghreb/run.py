#!/usr/bin/python

import os
import sys
import time
from scrapy import cmdline

name = 'LeadsScraper'
if __name__ == '__main__':
    #command = "scrapy crawl {0} -t csv -o YP_USA_Business_Outputs.csv".format(name).split()
    command = "scrapy crawl {0} -t csv -o Leads.csv".format(name).split()
    cmdline.execute(command)