# -*- coding: utf-8 -*-

import re
import time

from urlparse import urljoin
from subscene_v3.items import MovieItem
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request

import copy
import subscene_v3.settings
from scrapy import log
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher

class azre_scraper(Spider):

    name = "subscenev3_scraper"

    base_url = 'http://subscene.com/'

    inputs_urls = 'movies_subs_urls.txt'

    allowed_domains = ["subscene.com"]

    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_ids = set()

    inputs_urls_lst = list()

    def __init__(self):
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        self.inputs_urls_lst = subscene_v3.settings.loadFileIntoLIst(self.inputs_urls)
        subscene_v3.settings.remaining_urls = copy.deepcopy(self.inputs_urls_lst)

    def spider_closed(spider, reason):
        log.msg('Reason to close the spider:' + reason)
        rem = subscene_v3.settings.remaining_urls
        f = open('movies_subs_urls.txt', 'w')
        for r in rem:
            f.write(r + '\n')
        f.close()

    def start_requests(self):
        meta = dict()
        for url in self.inputs_urls_lst:
            meta['source_url'] = url
            yield Request(url=url,
                          meta=copy.deepcopy(meta),
                          headers=self.headers,
                          callback=self.parse_individual_page)

    def parse_individual_page(self, response):
        sel = Selector(response)

        imdb = sel.xpath('.//a[@class="imdb"]/@href')
        if imdb and len(imdb) > 0:
            imdb = imdb.extract()[0].replace('http://www.imdb.com/title/', '').strip()

            name = sel.xpath('.//h1/span[@itemprop="name"]/text()')
            if name and len(name) > 0:
                name = name.extract()[0].strip()
                name = re.sub('\r|\n', '', name).strip()

            ratings = ''
            rating_users = ''
            rating = sel.xpath('.//div[@class="rating positive hint hint--left"]')
            if rating and len(rating) > 0:
                rating_users = rating.xpath('@data-hint')
                if rating_users and len(rating_users) > 0:
                    rating_users = rating_users.extract()[0].replace('By', '').replace('users', '').strip()
                ratings = rating.xpath('.//span/text()')
                if ratings and len(ratings) > 0:
                    ratings = ratings.extract()[0].strip()

            hearing_impaired = sel.xpath('.//span[@class="hearing-impaired"]')
            if hearing_impaired and len(hearing_impaired):
                hearing_impaired = 'Yes'
            else:
                hearing_impaired = 'No'

            lang = re.findall('/([a-z]+)/\d+', response.url)
            if lang and len(lang) > 0:
                lang = lang[0].strip()
            else:
                lang = ''

            author = sel.xpath('.//li[@class="author"]/a')
            if author and len(author):
                author_name = author.xpath('text()').extract()[0].strip()
                author_url = author.xpath('@href').extract()[0].strip()
                author_url = urljoin(self.base_url, author_url)
            else:
                author_name = ''
                author_url = ''

            subtitle_url = sel.xpath('//*[@id="downloadButton"]/@href')
            if subtitle_url and len(subtitle_url) > 0:
                subtitle_url = subtitle_url.extract()[0]
                subtitle_url = urljoin(self.base_url, subtitle_url)
            source_url = response.meta['source_url']

            item = MovieItem()
            item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
            item['Movie_IMDB_ID'] = imdb
            item['Movie_Name'] = name
            item['SubtitleURL'] = subtitle_url
            item['Rating'] = ratings
            item['RatingUsersCount'] = rating_users
            item['HI'] = hearing_impaired
            item['Lang'] = lang
            item['Owner'] = author_name
            item['OwnerURL'] = author_url
            item['ListingURL'] = response.url
            item['SourceURL'] = source_url

            yield item