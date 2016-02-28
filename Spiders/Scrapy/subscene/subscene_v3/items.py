# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy

class MovieItem(scrapy.Item):
    TimeStamp = scrapy.Field()
    Movie_IMDB_ID = scrapy.Field()
    Movie_Name = scrapy.Field()
    Rating = scrapy.Field()
    RatingUsersCount = scrapy.Field()
    HI = scrapy.Field()
    Lang = scrapy.Field()
    Owner = scrapy.Field()
    OwnerURL = scrapy.Field()
    ListingURL = scrapy.Field()
    SubtitleURL = scrapy.Field()
    SourceURL = scrapy.Field()