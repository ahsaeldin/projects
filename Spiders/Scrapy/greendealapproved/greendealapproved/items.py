# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# http://doc.scrapy.org/en/latest/topics/items.html

import scrapy
from scrapy import Field

class GreendealapprovedItem(scrapy.Item):
    Name = Field()
    CompanyName = Field()
    CertificationID = Field()
    DateCertified = Field()
    Telephone = Field()
    Website = Field()
    Email = Field()
    Address = Field()
    DateExtracted = Field()
    GreenDealServicesOffered = Field()
    AdvertURL = Field()