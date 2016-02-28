# -*- coding: utf-8 -*-
import re

from scrapy.contrib.exporter import CsvItemExporter


class FormatItemValuePipeline(object):
    """Pipeline for formatting an item's values"""

    def process_item(self, item, spider):

        for key, value in dict(item).items():
            item[key] = ''.join(value)

        return item


class CsvExportPipeline(object):
    """Pipeline for exporting the scraped items to a .csv file"""

    def __init__(self):
        self.file = None
        self.exporter = None

    def open_spider(self, spider):
        self.file = open('{spider_name}.csv'.format(spider_name=spider.name), 'wb')
        self.exporter = CsvItemExporter(self.file)
        self.exporter.start_exporting()

    def process_item(self, item, spider):
        self.exporter.export_item(item)
        return item

    def close_spider(self, spider):
        self.exporter.finish_exporting()
        self.file.close()