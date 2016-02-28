# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

from scrapy.item import Item, Field

class LeadItem(Item):
    TimeStamp = Field()
    Game_ID = Field()
    Name = Field()
    NoOfReleases = Field()
    Desc = Field()
    CoverArt = Field()
    GameDetails = Field()
    RefURL = Field()
    OverViewRef = Field()
