# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class MoxaItem(scrapy.Item):
    # define the fields for your item here like:

    SKU = scrapy.Field()
    MainCategory = scrapy.Field()
    Subcategory1 = scrapy.Field()
    Subcategory2 = scrapy.Field()
    Subcategory3 = scrapy.Field()
    Subcategory4 = scrapy.Field()
    Images = scrapy.Field()
    Datasheet = scrapy.Field()
    FeaturesAndBenefits = scrapy.Field()
    Overview = scrapy.Field()
    AvailableModels = scrapy.Field()
    CompatibleModules = scrapy.Field()
    HardwareSpecifications = scrapy.Field()
    Weight = scrapy.Field()
    Dimensions = scrapy.Field()
    RelatedProducts = scrapy.Field()

