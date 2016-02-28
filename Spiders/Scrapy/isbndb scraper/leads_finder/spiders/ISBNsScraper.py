# -*- coding: utf-8 -*-

import re
from scrapy.selector import HtmlXPathSelector
from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor

class BusinessListCrawler(CrawlSpider):

    download_delay = 5
    jobs_list = []
    output_file_name = 'isbns.txt'
    name = "isbns"
    #allowed_domains = ['http://isbndb.com']
    start_urls = ['http://www.amazon.com/books-used-books-textbooks/b?node=283155']

    rules = (
        Rule(SgmlLinkExtractor(allow=('\w+',)), callback='parse_page', follow=True),
    )

    seen_isbns = set()
    def parse_page(self, response):
        hxs = HtmlXPathSelector(response)

        isbns = re.findall('/dp/(\w{10})', response.body)
        if isbns and len(isbns) > 0:
            isbns = [x for x in isbns if len(x) == 10 and not x.startswith('B')]

        if len(isbns) > 0:
            isbns = list(set(isbns))
            f = open(self.output_file_name, 'ab')
            for link in isbns:
                if not link in self.seen_isbns:
                    f.write(link + '\r')
            f.close()