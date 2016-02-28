# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html

import os
import dblite
import unicodecsv as csv
from yp.items import YPItem
from yp.items import YPSqliteItem
from scrapy.exceptions import DropItem

class YPPipeline(object):

    outputs_file_name = 'YP__Organized.csv'

    def process_item(self, item, spider):
        file_exists = False
        if os.path.exists(self.outputs_file_name):
            file_exists = True

        headers = ['TimeStamp',
                   'ZipCode',
                   'Keyword',
                   'Name',
                   'Phone',
                   'Categories',
                   'Website',
                   'Address',
                   'ListingURL',
                   'OverviewPageURL']

        test_file = open(self.outputs_file_name, 'ab')
        csvwriter = csv.DictWriter(test_file, quotechar='"', delimiter=',', fieldnames=headers)
        if not file_exists:
            csvwriter.writeheader()
        csvwriter.writerow(item)
        test_file.close()

        return item


#More info:https://pypi.python.org/pypi/scrapy-dblite/0.2.5
#http://stackoverflow.com/questions/3261858/does-anyone-have-example-code-for-a-sqlite-pipeline-in-scrapy
class YPSQLitePipeline(object):
    def __init__(self):
        self.ds = None

    def open_spider(self, spider):
        #self.ds = dblite.open(YPSqliteItem, 'sqlite://YellowPageSSsUS.db:YP_2', autocommit=True)
        self.ds = dblite.open(YPItem, 'sqlite://YP.db:YP_2', autocommit=True)

    def process_item(self, item, spider):
        if isinstance(item, YPItem):
            try:
                self.ds.put(item)
            except dblite.DuplicateItem:
                for p in self.ds.get({'Phone': str(item['Phone'])}):
                    zip_code = p['ZipCode']
                    if zip_code == '':
                        p['ZipCode'] = item['ZipCode']
                    p['Address'] = item['Address']
                    p['Website'] = item['Website']
                    p['Categories'] = item['Categories']
                    self.ds.put(p)
                #raise DropItem("Duplicate item found: %s" % item)
        else:
            raise DropItem("Unknown item type, %s" % type(item))
        return item

    def close_spider(self, spider):
        self.ds.commit()
        self.ds.close()