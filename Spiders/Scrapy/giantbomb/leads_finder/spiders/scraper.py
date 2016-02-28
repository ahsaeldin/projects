# -*- coding: utf-8 -*-

import re
import time
from scrapy import log
import unicodecsv as csv
from urlparse import urljoin
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request
from leads_finder.items import LeadItem
from scrapy.http import FormRequest, Request

class LeadsScraper(Spider):

    name = "games_scraper"
    base_url = 'http://www.giantbomb.com/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_listings_per_session = 0

    def start_requests(self):

        test_file = 'games_urls.csv'
        csv_file = csv.DictReader(open(test_file, 'rb'), delimiter=',')
        for line in csv_file:
            game_url = line['game_url']
            overview = line['overview']

            game_url = game_url.replace('/%0A>', '').strip()

            #cached_game_url = 'http://webcache.googleusercontent.com/search?q=cache:' + game_url.replace('http://', '')

            cached_game_url = game_url #.replace('http://', '')

            meta = dict()

            self.headers['Referer'] = overview

            meta['overview'] = overview.strip()

            yield Request(url=cached_game_url,
                          headers=self.headers,
                          meta=meta,
                          callback=self.parse_game,
                          errback=self.err_download)

    def err_download(self, failure):
        url = failure.request.url
        f = open('failed_urls', 'ab')
        f.write(url + '\r')
        f.close()

    def parse_game(self, response):
        self.log_line(response.url)
        sel = Selector(response)

        overview = response.meta['overview']

        listing_id = re.findall('data-object-guid="([\d\-]+)"', response.body)
        if listing_id and len(listing_id) > 0:
            listing_id = listing_id[0].strip()
        else:
            listing_id = ''
            self.log_line('Cannot find listing id::' + response.url)

        name = sel.xpath('.//h1/a[@class="wiki-title"]/text()')
        if name and len(name) > 0:
            name = name.extract()[0].strip()
        else:
            name = ''
            self.log_line('Cannot find listing name::' + response.url)

        desc = sel.xpath('.//div[@data-field="description"]/descendant::p/text()')
        if desc and len(desc) > 0:
            desc = desc.extract()
            desc = [x.strip() for x in desc]
            desc = [x for x in desc if x != '']
            desc = ' '.join(desc).replace(',', ';')
            desc = re.sub('\r|\n', ' ', desc)
        else:
            desc = ''
            self.log_line('Cannot find listing desc::' + response.url)

        image = sel.xpath('.//div[@class="wiki-boxart imgboxart imgcast"]/img/@src')
        if image and len(image) > 0:
            image = image.extract()[0].strip()
        else:
            image = ''
            self.log_line('Cannot find listing image::' + response.url)

        details = {}
        details_table = sel.xpath('.//div[@class="wiki-details"]/table/tbody/tr')
        if details_table and len(details_table) > 0:
            for detail in details_table:
                detail_name = detail.xpath('th/text()')
                if detail_name and len(detail_name) > 0:
                    detail_name = detail_name.extract()[0].strip()
                    detail_value = detail.xpath('td/.//*/text()')
                    if detail_value and len(detail_value) > 0:
                        detail_value = detail_value.extract()
                        detail_value = [x.strip() for x in detail_value]
                        detail_value = [x for x in detail_value if x != '']
                        if len(detail_value) > 1:
                            details[detail_name] = ' | '.join(detail_value)
                        elif len(detail_value) == 1:
                            details[detail_name] = detail_value[0].strip()

        releases_no = sel.xpath('.//*[@class="wiki-descriptor"]/a/text()')
        if releases_no and len(releases_no) > 0:
            releases_no = releases_no.extract()[0].strip()
            releases_no = releases_no.replace('releases', '').strip()
        else:
            releases_no = ''
            self.log_line('Cannot find no of releases::' + response.url)

        item = LeadItem()
        item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
        item['Game_ID'] = listing_id
        item['Name'] = name
        item['NoOfReleases'] = releases_no
        item['Desc'] = desc
        item['CoverArt'] = image
        item['GameDetails'] = details
        item['RefURL'] = response.url.replace('http://webcache.googleusercontent.com/search?q=cache:', 'http://')
        item['OverViewRef'] = overview

        # Only if there is an image.
        #if image != '':
        yield item

    def log_line(self, ms):
        disabled = False
        log.msg(ms, log.INFO)
        # print('FYI:' + ms)
        # f = open('log.txt', 'ab')
        # f.write(ms + '\n')
        # f.close()
