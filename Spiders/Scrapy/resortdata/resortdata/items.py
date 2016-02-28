# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

from scrapy.item import Item, Field

class ResortdataItem(Item):
    # define the fields for your item here like:
    # name = Field()
    h1 = Field()
    banner_image_url = Field()
    body_stripped = Field()
    url = Field()
