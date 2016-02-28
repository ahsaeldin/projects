# -*- coding: utf-8 -*-

import re
import time
import copy
from yp.items import YPItem
from urlparse import urljoin
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request

class yp_us_scraper(Spider):

    name = "yp_us_scraper"
    keywords = 'keywords.txt'
    zip_codes = 'zip_codes.txt'
    done_zip_codes = 'done_zipcodes.txt'
    allowed_domains = ['yellowpages.com']
    base_url = 'http://www.yellowpages.com'
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

        z_file = open(self.zip_codes)
        zip_codes = z_file.read().splitlines()
        zip_codes = list(set(zip_codes))
        z_file.close()

        d_file = open(self.done_zip_codes)
        done_zip_codes_list = d_file.read().splitlines()
        done_zip_codes_list = list(set(done_zip_codes_list))
        d_file.close()

        zip_codes = [x for x in zip_codes if x not in done_zip_codes_list]

        zip_codes_count = len(zip_codes)
        keywords_count = len(keywords_list)

        zip_codes_counter = 0
        for zip_code in zip_codes:
            meta = dict()

            keywords_counter = 0
            zip_codes_counter += 1
            for keyword in keywords_list:
                keywords_counter += 1
                search_term = re.sub('&|\(|\)|/', ' ', keyword.lower()).replace(' ', '+')
                search_term = re.sub('\+{2,}', '+', search_term)
                url = 'http://www.yellowpages.com/search?search_terms=%s&geo_location_terms=%s' % (search_term.strip() , zip_code.strip())

                meta['keyword'] = keyword
                meta['zip_code'] = zip_code

                print ('Locations: ' + str(zip_codes_counter) + ' / ' + str(zip_codes_count)) + ' -- ' + \
                       str(keywords_counter) + '/' + str(keywords_count) + ' Extracted >> ' + \
                       str(self.seen_listings_per_session) + ' << YAY'

                yield Request(url=url,
                              meta=copy.deepcopy(meta),
                              headers=self.headers,
                              callback=self.parse_overview_page)

            with open(self.done_zip_codes, 'ab') as f:
                f.write(zip_code.strip() + '\r\n')

    def parse_overview_page(self, response):
        sel = Selector(response)

        listings = sel.xpath('.//div[@class="v-card"]')
        if listings and len(listings) > 0:
            for listing in listings:
                phone_no = listing.re('class="phones phone primary">(.+?)</div>')
                if phone_no and len(phone_no) > 0:
                    phone_no = phone_no[0].strip()

                    self.seen_listings_per_session += 1
                    name_node = listing.xpath('.//a[@itemprop="name"]')
                    if name_node and len(name_node) > 0:
                        name = name_node.xpath('text()').extract()[0].strip()
                        url = name_node.xpath('@href').extract()[0].strip()
                        url = urljoin(self.base_url, url).strip()

                        lnk_href = ''
                        related_links = listing.xpath('.//div[@class="links"]')
                        if related_links and len(related_links) > 0:
                            related_links = related_links.xpath('.//a')
                            for r in related_links:
                                lnk_text = r.xpath('text()')
                                if lnk_text and len(lnk_text) > 0:
                                    lnk_text = lnk_text.extract()[0].strip()
                                    if lnk_text == 'Website':
                                        lnk_href = r.xpath('@href')
                                        if lnk_href and len(lnk_href) > 0:
                                            lnk_href = lnk_href.extract()[0].strip()
                                            break

                        address = ''
                        address_node = listing.xpath('.//p[@itemprop="address"]')
                        if address_node and len(address_node) > 0:
                            address = address_node.xpath('.//text()')
                            if address and len(address) > 0:
                                address = address.extract()
                                address = [x.strip() for x in address]
                                address = [x for x in address if x != '']
                                address = '-'.join(address)

                                zip_code_from_address = address_node.xpath('.//span[@itemprop="postalCode"]/text()')
                                if zip_code_from_address and len(zip_code_from_address) > 0:
                                    zip_code_from_address = zip_code_from_address.extract()[0].strip()

                        categories = ''
                        categories_node = listing.xpath('.//div[@class="categories"]')
                        if categories_node and len(categories_node) > 0:
                            categories_node = categories_node.xpath('.//text()')
                            if categories_node and len(categories_node) > 0:
                                categories_node = categories_node.extract()
                                categories_node = [x.strip() for x in categories_node]
                                categories = ' | '.join(categories_node)

                        # if categories == '':
                        #     print 'asdf'

                        item = YPItem()
                        item['Name'] = copy.deepcopy(name)
                        item['Phone'] = copy.deepcopy(phone_no)
                        item['ListingURL'] = copy.deepcopy(url)
                        item['OverviewPageURL'] = response.url
                        item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
                        item['ZipCode'] = response.meta['zip_code']
                        item['Keyword'] = response.meta['keyword']
                        item['Website'] = copy.deepcopy(lnk_href)
                        item['Address'] = copy.deepcopy(address)
                        item['Categories'] = copy.deepcopy(categories)

                        yield item

        next_url = sel.xpath('.//a[@class="next ajax-page"]/@href')
        if next_url and len(next_url) > 0:
            next_url = next_url.extract()[0]
            next_url = urljoin(self.base_url, next_url)
            yield Request(url=next_url,
                          meta=response.meta,
                          headers=self.headers,
                          callback=self.parse_overview_page)


