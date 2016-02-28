# -*- coding: utf-8 -*-

import re
from bs4 import BeautifulSoup
from scrapy.selector import HtmlXPathSelector
from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor

class MoxaCrawler(CrawlSpider):

    name = "moxa_crawler"

    jobs_list = []
    allowed_domains = ['moxa.com']
    output_file_name = 'eheheheheheheheheheheheheh2.txt'
    start_urls = ['http://www.moxa.com/product/product.aspx']

    rules = (
        Rule(SgmlLinkExtractor(allow=('/product/',)), callback='parse_page', follow=True),
        # Rule(SgmlLinkExtractor(allow=('/location/',)), callback='parse_page', follow=True),
        # Rule(SgmlLinkExtractor(allow=('/companies/',)), callback='parse_page', follow=True),
        # Rule(SgmlLinkExtractor(allow=('/companies/.+?/\d+',)), callback='parse_page', follow=True),
    )

    def parse_page(self, response):
        links = []
        hxs = HtmlXPathSelector(response)
        for sel in hxs.xpath('//a'):
            link = sel.xpath('./@href').extract()
            for itm in link:
                if '/product/' in itm:
                    last_chr = itm[-1]
                    if last_chr == '/':
                        itm = itm[:-1]
                    itm_link = 'http://www.moxa.com' + itm.strip()
                    if '-' in itm_link or '_' in itm_link:
                        print itm_link
                        links.append(itm_link)

        if len(links) > 0:
            links = list(set(links))
            f = open(self.output_file_name, 'ab')
            for link in links:
                f.write(link + '\r')
            f.close()