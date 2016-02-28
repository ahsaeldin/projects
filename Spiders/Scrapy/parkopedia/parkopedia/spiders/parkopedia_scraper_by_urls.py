# -*- coding: utf-8 -*-

import os
import re
import time
import simplejson
from urlparse import urljoin
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request
from parkopedia.items import ParkoPediaItem

import copy
import parkopedia.settings
from scrapy import log
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher

class azre_scraper(Spider):

    name = "parkopedia_scraper"

    cities = 'cities.txt'

    not_found = 'Not_Found.txt'

    base_url = 'http://en.parkopedia.com.br/'

    allowed_domains = ["en.parkopedia.com.br"]

    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_ids = set()

    cities_lst = list()

    done_cities_count = 0

    download_delay = 10

    def __init__(self):
        if os.path.exists(self.not_found):
            os.remove(self.not_found)
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        self.cities_lst = parkopedia.settings.loadFileIntoLIst(self.cities)
        self.cities_lst = list(set(self.cities_lst))
        parkopedia.settings.remaining_cities = copy.deepcopy(self.cities_lst)

    def spider_closed(spider, reason):
        log.msg('Reason to close the spider:' + reason)
        rem = parkopedia.settings.remaining_cities
        f = open('cities.txt', 'w')
        for r in rem:
            f.write(r.strip() + '\n')
        f.close()

    def start_requests(self):
        inputs_urls_lst = list()
        cities = self.cities_lst
        for city in cities:
            url = 'http://en.parkopedia.com.br/parking/%s/' % city.strip().replace(' ', '_')
            meta = dict()
            meta['city'] = ''
            yield Request(url=url.strip(),
                          meta=copy.deepcopy(meta),
                          headers=self.headers,
                          callback=self.parse_city_page)

        # meta = dict()
        # meta['city'] = ''
        # yield Request(url='http://en.parkopedia.com.br/parking/carpark/hospital_santa_joana/52010/recife/',
        #               meta=copy.deepcopy(meta),
        #               headers=self.headers,
        #               callback=self.parse_city_page)


    def parse_city_page(self, response):
        self.done_cities_count += 1

        sel = Selector(response)

        not_found = sel.xpath('//*[@id="noresultsmessage"]')
        if not_found and len(not_found):
            not_found = not_found.extract()[0]
            if 'No parking spaces found' in not_found:
                file = open('Not_Found.txt', 'ab')
                file.write(response.meta['city'] + '\r\n')
                file.close()
                yield

        results = sel.xpath('//*[@id="search_results_data"]')
        rows = results.xpath('.//tr[starts-with(@id, "link")]')
        if rows and len(rows) > 0:

            parks_dicts = list()
            scripts = sel.xpath('.//script/text()')
            if scripts and len(scripts) > 0:
                for scr in scripts:
                    scr_lines = scr.extract()
                    if 'var _pspaces' in scr_lines:
                        scr_lines = re.sub('\r|\n', ' ', scr_lines).strip()
                        parks_json = re.findall('var\s+_pspaces\s=\s\[(.+?)\];', scr_lines)[0].strip()
                        parks_json = re.findall('{"id":"\d+".+?"currency":".+?"}]}', parks_json)
                        for p in parks_json:
                            p_dict = simplejson.loads(p.strip())
                            parks_dicts.append(p_dict)

            for row in rows:
                cells = row.xpath('.//td')

                park_id = ''
                park_url = ''
                address = ''
                distance = ''
                info = ''
                pricing = ''
                for cell in cells:
                    class_name = cell.xpath('@class')
                    if class_name and len(class_name) > 0:
                        class_name = class_name.extract()[0].strip()

                        if class_name == 'column col_key':
                            park_anchor = cell.xpath('.//a')
                            if park_anchor and len(park_anchor) > 0:
                                park_id = park_anchor.xpath('@id').extract()[0]
                                park_url = park_anchor.xpath('@href').extract()[0]
                                park_url = urljoin(self.base_url, park_url)
                        elif class_name == 'column col_address':
                            park_address_anchor = cell.xpath('.//a/text()')
                            if park_address_anchor and len(park_address_anchor) > 0:
                                address = park_address_anchor.extract()[0].strip()
                        elif class_name == 'column col_distance':
                            distance = cell.xpath('.//span/@value')
                            if distance and len(distance) > 0:
                                distance = distance.extract()[0].strip()
                        elif class_name == 'column col_info':
                            info = cell.xpath('text()')
                            if info and len(info) > 0:
                                info = info.extract()[0]
                        elif class_name == 'column col_pricing':
                            pricing = cell.xpath('text()')
                            if pricing and len(pricing) > 0:
                                pricing = pricing.extract()[0]

                type_str = ''
                features_str = ''
                accepted_payments = ''
                sidebar = sel.xpath('//*[@id="sidebarcontent"]')
                if sidebar and len(sidebar) > 0:
                    features = sidebar.xpath('.//div[@class="spaceinfocontainer"]')
                    if features and len(features) > 0:
                        features = features.xpath('.//p/text()')
                        if features and len(features):
                            features = features.extract()
                            for f in features:
                                if 'Features' in f:
                                    features_str = f.replace('Features:', '').strip()
                                elif 'Type' in f:
                                    type_str = f.replace('Type:', '').strip()
                                elif 'Accepted Payments' in f:
                                    accepted_payments = f.replace('Accepted Payments:', '').strip()


                item = ParkoPediaItem()
                item['ParkID'] = park_id
                item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
                item['ParkURL'] = park_url
                item['Address'] = address
                item['Progress'] = self.done_cities_count

                item['Country'] = 'BR'
                item['State'] = ''
                item['ParkingType'] = type_str
                item['PaymentsAccepted'] = accepted_payments
                item['Type'] = type_str
                item['Features'] = features_str

                print response.url

                for park in parks_dicts:
                    park_id_json = copy.deepcopy(park['rid'])
                    if park_id == park_id_json:
                        item['ParkName'] = unicode(copy.deepcopy(park['co'])) if park.has_key('co') else ''
                        if item['ParkName'] == '':
                            item['ParkName'] = unicode(copy.deepcopy(park['title'])) if park.has_key('title') else ''
                            if item['ParkName'] == '':
                                item['ParkName'] == address
                        item['Phone'] = copy.deepcopy(park['ph']) if park.has_key('ph') else ''
                        item['NumberofParkingSpaces'] = copy.deepcopy(park['num']) if park.has_key('num') else ''
                        item['Latitude'] = copy.deepcopy(park['lat'])
                        item['Longitude'] = copy.deepcopy(park['lng'])
                        item['City'] = unicode(copy.deepcopy(park['city']))
                        item['Info'] = park['hpschema']
                        item['Neighborhood'] = unicode(copy.deepcopy(park['area']))
                        item['PostCode'] = copy.deepcopy(park['code'])
                        break

                if item['ParkName'] == '':
                    item['ParkName'] == address

                print park_url
                yield item

