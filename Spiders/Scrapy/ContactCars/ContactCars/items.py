# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

from scrapy.item import Item, Field

class ContactcarsItem(Item):
    ScrapeTime = Field()
    CC = Field()
    Make = Field()
    Model = Field()
    Year = Field()
    Price = Field()
    URL = Field()