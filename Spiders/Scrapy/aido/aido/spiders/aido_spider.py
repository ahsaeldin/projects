# -*- coding: utf-8 -*-

import re
import time
import copy
import requests
from scrapy import log
from urlparse import urljoin
from aido.items import AidoItem
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request
from scrapy.selector import HtmlXPathSelector

class FoodSpider(Spider):

    name = "aido"
    categories_dict = dict()
    allowed_domains = ['aido.com']
    base_url = 'http://www.aido.com/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    def start_requests(self):
        yield Request(url='http://www.aido.com/eshop/faces/jsp/static.jsp?articleid=2018',
                      headers=self.headers,
                      callback=self.parse_site_map)

    def parse_site_map(self, response):
        sel = Selector(response)

        content = sel.xpath('//*[@id="staticArticleContent"]')
        if content and len(content):
            categories_links = content.xpath('.//a')
            if categories_links and len(categories_links) > 0:
                for categories_link in categories_links:
                    category_url = categories_link.xpath('@href').extract()[0]
                    category_name = categories_link.xpath('text()').extract()[0]
                    category_url = urljoin(self.base_url, category_url)
                    if '.html' in category_url:
                        self.categories_dict[category_url] = category_name
        else:
            log.msg('Error in the sitemap', log.INFO)

        for category_page_url, cat_name in self.categories_dict.iteritems():
              base_meta = dict()
              base_meta['category_name'] = cat_name
              meta = copy.deepcopy(base_meta)
              yield Request(url=category_page_url,
                            meta=meta,
                            headers=self.headers,
                            callback=self.pare_category_overview_page)

        print ';asdf'

    def pare_category_overview_page(self, response):
        sel = Selector(response)

        next_page = sel.xpath('.//li[@class="pagination-li-filed next-arrow"]')
        if next_page and len(next_page) > 0:
            next_page = next_page
            print response.url

#pagination-li-filed next-arrow
    # def parse_category_page(self, response):
    #     sel = HtmlXPathSelector(text=response.body)
    #
    #     links = sel.xpath('//a[starts-with(@id, "businessId")]/@href').extract()
    #
    #     for link in links:
    #         print link
    #         yield Request(
    #             url=link,
    #             meta=response.meta,
    #             headers=self.headers,
    #             callback=self.parse_prod_page
    #         )
    #
    #     next_page = sel.xpath('//a[starts-with(@onclick, "sendFormWithOptionsCU(")]/@href').extract()
    #     nextpage = next_page[-1]
    #     yield Request(
    #                 url=nextpage,
    #                 dont_filter=True,
    #                 meta=response.meta,
    #                 headers=self.headers,
    #                 callback=self.parse_category_page
    #     )
    #
    # def start_requests(self):
    #
    #     ddd = FoodItem()
    #     with open('categories.txt') as f:
    #         urls = f.read().splitlines()
    #
    #     for url in urls:
    #         base_meta = dict()
    #         base_meta['dont_merge_cookies'] = True
    #         meta = copy.deepcopy(base_meta)
    #
    #         yield Request(
    #             url=url,
    #             meta=meta,
    #             headers=self.headers,
    #             callback=self.parse_category_page
    #         )
    #
    # def parse_prod_page(self, response):
    #     sel = Selector(response)
    #     name = sel.xpath('//h1[@id="businessTitle"]/text()')
    #     if name:
    #         name = name.extract()
    #         name = name[0]
    #
    #         address = sel.xpath('//p[@itemprop="address"]/span/text()')
    #         if address:
    #             address = address.extract()
    #             address = '|'.join(address)
    #
    #         telephone = sel.xpath('//span[@itemprop="telephone"]/text()')
    #         if telephone:
    #             telephone = '|'.join(telephone.extract())
    #
    #         emails = ''
    #         website = sel.xpath('//dl[@class="m-business-card--online"]/dd/a/@href')
    #         if website:
    #             website = website.extract()[0]
    #             print 'Checking %s for emails...' % website
    #             req = requests.get(website)
    #             if req:
    #                 page_source = re.sub('\r|\n|\t', ' ', req.text).strip()
    #                 emails_matcher = re.findall('\w+@\w+\.\w+', page_source)
    #                 if emails_matcher:
    #                     emails = set()
    #                     for email in emails_matcher:
    #                         emails.add(email)
    #                     emails = list(emails)
    #                     emails = '|'.join(emails)
    #
    #         item = PaginasamarillasItem()
    #         item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
    #
    #         item['CompanyName'] = name
    #         item['Phone'] = telephone
    #         item['Website'] = website
    #         item['Address'] = address
    #         item['Email'] = emails
    #
    #         item['RefUrl'] = response.url
    #
    #         return item