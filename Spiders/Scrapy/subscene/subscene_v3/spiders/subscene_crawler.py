# -*- coding: utf-8 -*-

import re
from bs4 import BeautifulSoup
from scrapy.selector import Selector
from scrapy.selector import HtmlXPathSelector
from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor

class SubsceneCrawler(CrawlSpider):

    jobs_list = []
    output_file_name = 'Subs_URLs.txt'
    name = "SubsceneCrawler"
    allowed_domains = ['subscene.com']
    start_urls = ['http://subscene.com/']

    rules = (
        Rule(SgmlLinkExtractor(allow=('\w+',),
                               deny=('subtitle/download', 'login', 'subscene.com/subtitles/.+?/[a-z]+/\d+'),),
             callback='parse_page', follow=True),
    )

    def parse_page(self, response):
        links_set = set()
        sel = Selector(response)
        links = sel.xpath('.//a/@href')
        if links and len(links) > 0:
            for link in links:
                xlink = link.extract()
                if 'http://subscene.com/subtitles/' in xlink:
                    links_set.add(xlink)
                    print xlink
                # if '/company/' in itm and not '#' in itm:
                #     last_chr = itm[-1]
                #     if last_chr == '/':
                #         itm = itm[:-1]
                #     itm_link = 'http://www.businesslist.net.nz' + itm.strip()
                #     links.append(itm_link)

        links_set = list(links_set)
        if len(links_set) > 0:
            f = open(self.output_file_name, 'ab')
            for link2 in links_set:
                f.write(link2 + '\n')
            f.close()