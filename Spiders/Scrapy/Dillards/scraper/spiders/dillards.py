# -*- coding: utf-8 -*-
import re
import time
import scrapy
import urlparse
from scrapy import log
from scrapy.spider import Spider
from scraper.items import ScraperItem
from scrapy.http.request import Request

class DillardsSpider(Spider):

    log.msg("Spider for Dillards.com started", level=log.INFO)
    name = "dillards"
    allowed_domains = ["dillards.com", ]

    start_urls = [
        'http://www.dillards.com/webapp/wcs/stores/servlet/ContextChooserView?storeId=301&catalogId=301&langId=-1',
    ]

    def parse(self, response):
        # Because they obtain the link through javascript, we need to read the country short name from the
        # a javascript file
        jsp_url = response.xpath("//script[@src='/javascript/fiftyOne.js']/@src").extract()
        url = urlparse.urljoin(response.url, jsp_url[0])
        yield Request(url, callback=self.parse_jsp)

    def parse_jsp(self, response):
        # extracting data from the jsp file
        country_data = []

        log.msg("Looking for all the countries flags", level=log.INFO)

        for data in re.findall('countryCodeArray\[\d{1,3}\]="(.+)\"', response.body):
            country_data.append(data.split("|"))

        log.msg("Found {} countries".format(len(country_data)))

        request = Request(self.start_urls[0], callback=self.parse_countries_page, dont_filter=True)
        request.meta['country_data'] = country_data
        yield request

    def parse_countries_page(self, response):

        countries = response.meta['country_data']

        for c in countries:
            # http://www.dillards.com/
            # country=AG; currency=USD
            # Because the name of the country doesn't appear on the home page we need to extract it from here
            item = ScraperItem()
            item['country'] = c[0]

            request_with_cookies = Request(
                url="http://www.dillards.com/",
                cookies={'currency': c[1], 'country': c[2]},
                callback=self.parse_home_page,
                dont_filter=True
            )
            request_with_cookies.meta['item'] = item
            yield request_with_cookies

    def parse_home_page(self, response):
        item = ScraperItem(response.meta['item'])
        f = open(item['country'] + '.html', 'ab')
        f.write(response.body)
        f.close()
        item['page_title'] = response.xpath("//title/text()").extract()
        yield item
