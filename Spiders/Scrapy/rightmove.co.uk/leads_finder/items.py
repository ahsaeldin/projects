# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

from scrapy.item import Item, Field

class LeadItem(Item):
    TimeStamp = Field()
    Listing_ID = Field()
    Address = Field()
    # Address0 = Field()
    # Address1 = Field()
    # Address2 = Field()
    # Address3 = Field()
    # Address4 = Field()
    # Address5 = Field()
    # Address6 = Field()
    PropertyValue = Field()
    ListingURL = Field()
    BrochureURL = Field()
    EPC_RatingURL = Field()
    ListingAgent = Field()