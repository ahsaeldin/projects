# -*- coding: utf-8 -*-

import re
import time
import copy
from scrapy.log import log
from urlparse import urljoin
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request
from leads_finder.items import LeadItem
from scrapy.http import FormRequest, Request

class LeadsScraper(Spider):

    name = "LeadsScraper"
    keywords = 'input_keywords.txt'
    allowed_domains = ['rightmove.co.uk']
    base_url = 'http://www.rightmove.co.uk/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_listings_per_session = 0

    download_delay = 3

    def start_requests(self):

        file = open(self.keywords)
        keywords_list = file.read().splitlines()
        keywords_list = list(set(keywords_list))
        file.close()

        keywords_count = str(len(keywords_list))

        keywords_counter = 0
        for keyword in keywords_list:

            keywords_counter += 1
            search_term = re.sub('&|\(|\)|/', ' ', keyword.lower()).replace(' ', '+')
            search_term = re.sub('\+{2,}', '+', search_term)

            url = 'http://www.rightmove.co.uk/property-for-sale/search.html?searchLocation=%s' % search_term

            meta = dict()
            meta['keyword'] = keyword
            meta['keywords_count'] = keywords_count
            meta['keywords_counter'] = keywords_counter

            yield Request(url=url,
                          meta=meta,
                          headers=self.headers,
                          callback=self.request_form_page)

    def request_form_page(self, response):
        sel = Selector(response)
        keyword = response.meta['keyword']
        search_form = sel.xpath('.//*[@id="propertySearchCriteria"]')
        if not search_form or search_form == 0:
            self.log_line('Search term "%s" not found.' % keyword)
        else:
            yield FormRequest.from_response(response,
                                            meta=response.meta,
                                            formname='propertySearchCriteria',
                                            callback=self.parse_overview_page)

    def parse_overview_page(self, response):

        page_no = re.findall('index=(\d+)', response.url)
        if page_no and len(page_no) > 0:
            page_no = str((int(page_no[0]) / 10 + 1))
        else:
            page_no = str(1)

        keyword = response.meta['keyword']
        keywords_count = response.meta['keywords_count']
        keywords_counter = response.meta['keywords_counter']

        body = response.body
        if  'The page you are attempting to access is unable to be served' in body:
            self.log_line('Keyword: %s - page # %s, Bot detected by server' % (keyword, page_no))

        self.log_line('Keyword: %s - %s' % (keyword, str(keywords_counter)) + '/' + str(keywords_count)
                      + ' - Page # ' + page_no
                      + ' - Extracted Listings Count >> ' + str(self.seen_listings_per_session))

        sel = Selector(response)

        listings = sel.xpath('.//div[@class="price-new touchsearch-summary-list-item-price"]/a/@href')
        if listings and len(listings) > 0:
            for listing in listings:
                listing_url = urljoin(response.url, listing.extract())
                yield Request(url=listing_url,
                              meta=response.meta,
                              headers=self.headers,
                              callback=self.parse_listing_page)
        else:
            self.log_line('Cannot find listings:' + response.url)

        next_url = sel.xpath('.//a[@class="pagenavigation active"]')
        if next_url and len(next_url) > 0:
            for n in next_url:
                next_text = n.xpath('text()').extract()[0]
                if next_text == 'next':
                    next_href = n.xpath('@href').extract()[0]
                    next_href = urljoin(self.base_url, next_href)
                    yield Request(url=next_href,
                                  meta=response.meta,
                                  headers=self.headers,
                                  callback=self.parse_overview_page)
        else:
            self.log_line('Cannot find next url:' + response.url)

    def parse_listing_page(self, response):

        self.seen_listings_per_session += 1

        sel = Selector(response)

        listing_id = re.findall('property\-(\d+)\.html', response.url)
        if listing_id and len(listing_id) > 0:
            listing_id = listing_id[0].strip()
        else:
            self.log_line('Cannot find listing id::' + response.url)

        property_value = ''
        price = sel.xpath('.//*[@id="propertyHeaderPrice"]/strong/text()')
        if price and len(price) > 0:
            property_value = price.extract()[0].strip()
            property_value = re.sub('\r|\n|\t', '', property_value)
            property_value = property_value.replace('|POA', '| POA')

        address_lst = []
        address = sel.xpath('.//address[@class="pad-0 fs-16 grid-25"]/text()')
        if address and len(address) > 0:
            address = address.extract()[0].strip()
            address = re.sub('\r\n', ' ', address)
            address_lst = address.split(',')
            last_item_in_address = address_lst[-1].strip()
            if len(last_item_in_address) == 3:
                address_lst = address_lst[:-1]
            address_lst = [x.strip() for x in address_lst]

        post_code = re.findall('propertyPostcode: "(.+?)",', response.body)
        if post_code and len(post_code) > 0:
            post_code = unicode(post_code[0].strip())
            address_lst = [x.replace(post_code, '') for x in address_lst]
            post_code_arr = post_code.split(' ')
            if not post_code in address_lst:
                address_lst.append(post_code)
            address_lst = [x.strip() for x in address_lst if not any(y == x for y in post_code_arr) and x != '']
        address_lst = ' | '.join(address_lst)

        brochure_url = sel.xpath('.//a[@id="brochure-1"]/@href')
        if brochure_url and len(brochure_url) > 0:
            brochure_url = urljoin(self.base_url, brochure_url.extract()[0])

        epc_rating_url = sel.xpath('.//a[@class="js-ga-hipepc"]/@href')
        if epc_rating_url and len(epc_rating_url) > 0:
            epc_rating_url = urljoin(self.base_url, epc_rating_url.extract()[0])

        listing_agent = sel.xpath('.//*[@id="aboutBranchLink"]/strong/text()')
        if listing_agent and len(listing_agent) > 0:
            listing_agent = listing_agent.extract()[0].strip()

        item = LeadItem()
        item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
        item['Listing_ID'] = listing_id
        item['PropertyValue'] = copy.deepcopy(property_value)
        item['Address'] = copy.deepcopy(address_lst)
        item['ListingURL'] = copy.deepcopy(response.url)
        item['BrochureURL'] = copy.deepcopy(brochure_url)
        item['EPC_RatingURL'] = copy.deepcopy(epc_rating_url)
        item['ListingAgent'] = copy.deepcopy(listing_agent)

        yield item

    def log_line(self, ms):
        print('FYI:' + ms)
        f = open('log.txt', 'ab')
        f.write(ms + '\n')
        f.close()
