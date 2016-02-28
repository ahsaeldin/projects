#!/usr/bin/python

from scrapy import cmdline

#ListsScraper
#ISBNsExtractor

name = 'ISBNsExtractor'
if __name__ == '__main__':
    #command = "scrapy crawl {0} -t csv -o YP_USA_Business_Outputs.csv".format(name).split()
    command = "scrapy crawl {0}".format(name).split()
    cmdline.execute(command)