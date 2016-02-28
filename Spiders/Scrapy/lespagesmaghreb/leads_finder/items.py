# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy

class LeadItem(scrapy.Item):
    Name = scrapy.Field()
    Phone = scrapy.Field()
    Phone = scrapy.Field()
    Address = scrapy.Field()
    OverviewPageURL = scrapy.Field()
    ListingID = scrapy.Field()