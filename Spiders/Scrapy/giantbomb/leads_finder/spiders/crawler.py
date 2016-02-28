# -*- coding: utf-8 -*-

import re
import time
from urlparse import urljoin
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request
from leads_finder.items import LeadItem
from scrapy.http import FormRequest, Request

class LeadsScraper(Spider):

    name = "games_crawler"
    base_url = 'http://www.giantbomb.com/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_listings_per_session = 0

    def start_requests(self):
        input_urls = {'nintendo-entertainment-system': 'http://www.giantbomb.com/nintendo-entertainment-system/3045-21/games/',
                      'xbox-one': 'http://www.giantbomb.com/xbox-one/3045-145/games/',
                      'playstation-4': 'http://www.giantbomb.com/playstation-4/3045-146/games'}

        urls_count = '3'

        urls_counter = 0
        for file_name, cat_url in input_urls.iteritems():

            urls_counter += 1

            meta = dict()
            meta['file_name'] = file_name
            meta['urls_count'] = urls_count
            meta['urls_counter'] = urls_counter

            yield Request(url=cat_url,
                          meta=meta,
                          headers=self.headers,
                          callback=self.parse_overview)

    def parse_overview(self, response):

        print response.url

        page_no = re.findall('index=(\d+)', response.url)
        if page_no and len(page_no) > 0:
            page_no = str((int(page_no[0]) / 10 + 1))
        else:
            page_no = str(1)

        sel = Selector(response)

        f = open('games_urls.csv', 'ab')
        listings = sel.xpath('.//*[@id="site"]/.//ul/li[@class="related-game"]')
        if listings and len(listings) > 0:
            for listing in listings:

                img_url = listing.xpath('.//img[@src]')
                if img_url and len(img_url) > 0:
                    img_url = img_url.extract()[0]
                    game_url = listing.xpath('a/@href')
                    if game_url and len(game_url) > 0:
                        game_url = game_url.extract()[0]
                        game_url = urljoin(self.base_url, game_url)
                        f.write(response.url + ',' + game_url + '\n')
        else:
            self.log_line('Cannot find listings:' + response.url)
        f.close()

        next_url = sel.xpath('.//ul[@class="paginate js-table-paginator"]/li[@class="skip next"]/a/@href')
        if next_url and len(next_url) > 0:
            next_url = next_url.extract()[0]
            next_url = urljoin(self.base_url, next_url)
            yield Request(url=next_url,
                          meta=response.meta,
                          headers=self.headers,
                          callback=self.parse_overview)
        else:
            self.log_line('Cannot find next url or no more:' + response.url)
