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

class LeadsScraper(Spider):

    name = "LeadsScraper"
    keywords = 'input_keywords.txt'
    inputs_zipcodes = 'inputs_zipcodes.txt'
    done_zip_codes = 'done_zipcodes.txt'
    allowed_domains = ['lespagesmaghreb.com']
    base_url = 'http://www.lespagesmaghreb.com/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_listings_per_session = 0

    def start_requests(self):
        file = open(self.keywords)
        keywords_list = file.read().splitlines()
        keywords_list = list(set(keywords_list))
        file.close()

        z_file = open(self.inputs_zipcodes)
        input_zipcodes = z_file.read().splitlines()
        input_zipcodes = list(set(input_zipcodes))
        z_file.close()

        d_file = open(self.done_zip_codes)
        done_zip_codes_list = d_file.read().splitlines()
        done_zip_codes_list = list(set(done_zip_codes_list))
        d_file.close()

        print 'Check done codes...'
        input_zipcodes = [x for x in input_zipcodes if x not in done_zip_codes_list]

        zip_codes_count = len(input_zipcodes)
        keywords_count = len(keywords_list)

        zip_codes_counter = 0
        print 'Starting...'
        for zip_code in input_zipcodes:
            meta = dict()

            keywords_counter = 0
            zip_codes_counter += 1
            for keyword in keywords_list:
                keywords_counter += 1
                search_term = re.sub('&|\(|\)|/', ' ', keyword.lower()).replace(' ', '+')
                search_term = re.sub('\+{2,}', '+', search_term)
                url = 'http://www.lespagesmaghreb.com/firms?commit=&search[conditions][ou]=%s&search[keywords]=%s' \
                      % (zip_code.strip(), search_term.strip())
                print url
                meta['keyword'] = keyword
                meta['zip_code'] = zip_code

                print ('Location: ' + str(zip_codes_counter) + ' / ' + str(zip_codes_count)) + ' -- Keyword:' + \
                       str(keywords_counter) + '/' + str(keywords_count) + ' -- Extracted >> ' + \
                       str(self.seen_listings_per_session) + ' << YAYe PO - AU'

                yield Request(url=url,
                              meta=copy.deepcopy(meta),
                              headers=self.headers,
                              callback=self.parse_overview_page)

            # with open(self.done_zip_codes, 'ab') as f:
            #     f.write(zip_code.strip() + '\r\n')

    def parse_overview_page(self, response):
        print '>>>>>>>>' + response.url
        sel = Selector(response)

        listings = sel.xpath('//*[@id="firms"]/li')
        if listings and len(listings) > 0:
            for listing in listings:

                email = ''
                website = ''
                listing_url = ''

                listing_id = listing.xpath('@id')
                if listing_id and len(listing_id) > 0:
                    listing_id = listing_id.extract()[0].replace('firm_', '').strip()
                    listing_id = str(listing_id)

                    name = listing.xpath('.//p[@class="title"]/text()')
                    if name and len(name) > 0:
                        name = name.extract()[0].strip()
                    else:
                        name = ''
                        #log.msg('Error cannot find name ' + response.url, log.INFO)


                    phone = listing.xpath('.//p[@class="contact-phones"]/img/@src')
                    if phone and len(phone) > 0:
                        phone = phone.extract()
                        phone = [urljoin(self.base_url, x) for x in phone]
                        phone = ' | '.join(phone)
                    else:
                        phone = ''
                        #print 'Error cannot find phone'

                    address = listing.xpath('.//p[@class="main-address"]/text()')
                    if address and len(address) > 0:
                        address = address.extract()[0].strip()
                    else:
                        address = listing.xpath('.//p[@class="listing-address mappable-address"]/text()')
                        if address and len(address) > 0:
                            address = address.extract()[0].strip()
                        else:
                            address = ''
                            #print 'Error cannot find address'

                    item = LeadItem()
                    item['Name'] = copy.deepcopy(name)
                    item['ListingID'] = copy.deepcopy(listing_id)
                    item['Phone'] = copy.deepcopy(phone)
                    item['OverviewPageURL'] = response.url
                    item['Address'] = copy.deepcopy(address)

                    yield item

                else:

                    print 'ID Not Found, Plz Check'

        next_url = sel.xpath('.//a[@class="next_page"]/@href')
        if next_url and len(next_url) > 0:
            next_url = next_url.extract()[0]
            next_url = urljoin(self.base_url, next_url)
            yield Request(url=next_url,
                          meta=response.meta,
                          headers=self.headers,
                          callback=self.parse_overview_page)


