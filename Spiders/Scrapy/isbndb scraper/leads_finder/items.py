# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy

class LeadItem(scrapy.Item):
    TimeStamp = scrapy.Field()
    Book_ID = scrapy.Field()
    Amazon_Best_Sellers_Rank = scrapy.Field()
    ISBN_10 = scrapy.Field()
    ISBN_13 = scrapy.Field()
    ASIN = scrapy.Field()
    Amazon_title = scrapy.Field()
    Edition = scrapy.Field()
    cover_type = scrapy.Field()
    Amazon_author = scrapy.Field()
    pages_length = scrapy.Field()
    Amazon_publisher = scrapy.Field()
    publication_date = scrapy.Field()
    publication_year = scrapy.Field()
    shipping_weight = scrapy.Field()
    Amazon_language = scrapy.Field()
    is_excluded = scrapy.Field()
    exclude_reason = scrapy.Field()
