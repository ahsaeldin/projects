# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy

class ParkoPediaItem(scrapy.Item):
    ParkID = scrapy.Field()
    ParkURL = scrapy.Field()
    Address = scrapy.Field()
    Info = scrapy.Field()
    Pricing = scrapy.Field()
    ParkName = scrapy.Field()
    Phone = scrapy.Field()
    Country = scrapy.Field()
    State = scrapy.Field()
    City = scrapy.Field()
    ParkingType = scrapy.Field()
    PaymentsAccepted = scrapy.Field()
    Type = scrapy.Field()
    NumberofParkingSpaces = scrapy.Field()
    Features = scrapy.Field()
    TimeStamp = scrapy.Field()
    Progress = scrapy.Field()
    Latitude = scrapy.Field()
    Longitude = scrapy.Field()
    Neighborhood = scrapy.Field()
    PostCode = scrapy.Field()