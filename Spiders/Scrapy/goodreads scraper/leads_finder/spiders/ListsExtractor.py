# -*- coding: utf-8 -*-

import re
import time
import copy
from scrapy.log import log
from urlparse import urljoin
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request

class LeadsScraper(Spider):

    name = "ListsScraper"
    input_urls = 'lists_input_urls.txt'
    done_zip_codes = 'done_input_urls.txt'
    allowed_domains = ['goodreads.com']
    base_url = 'http://www.goodreads.com/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_listings_per_session = 0
    seen_list = list()

    def start_requests(self):
        z_file = open(self.input_urls)
        input_zipcodes = z_file.read().splitlines()
        input_zipcodes = list(set(input_zipcodes))
        z_file.close()

        # d_file = open(self.done_zip_codes)
        # done_zip_codes_list = d_file.read().splitlines()
        # done_zip_codes_list = list(set(done_zip_codes_list))
        # d_file.close()
        #
        # print 'Check done codes...'
        # input_zipcodes = [x for x in input_zipcodes if x not in done_zip_codes_list]

        zip_codes_counter = 0
        print 'Starting...'
        for zip_code in input_zipcodes:
            zip_codes_counter += 1
            meta = dict()
            meta['counter'] = str(zip_codes_counter)
            meta['next_page_no'] = '1'
            yield Request(url=zip_code,
                          meta=copy.deepcopy(meta),
                          headers=self.headers,
                          callback=self.parse_overview_page)

    def parse_overview_page(self, response):
        print response.url
        sel = Selector(response)

        lists = sel.xpath('.//a[@class="listTitle"]/@href')
        if lists and len(lists) > 0:
            for list in lists:
                list_url = list.extract()
                list_url = urljoin(self.base_url, list_url)
                if not list_url in self.seen_list:
                    f = open('LISTs.txt', 'ab')
                    f.write(list_url + '\n')
                    f.close()

        next_url = sel.xpath('.//a[@class="next_page"]/@href')
        if next_url and len(next_url) > 0:
            next_url = next_url.extract()[0]
            next_url = urljoin(self.base_url, next_url)
            next_page_no = re.findall('page=(\d+)', next_url)
            if next_page_no and len(next_page_no) > 0:
                next_page_no = next_page_no[0]
                response.meta['next_page_no'] = next_page_no
            yield Request(url=next_url,
                          meta=response.meta,
                          headers=self.headers,
                          callback=self.parse_overview_page)