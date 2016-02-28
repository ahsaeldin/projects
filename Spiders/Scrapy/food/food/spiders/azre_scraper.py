# -*- coding: utf-8 -*-

import re
import time

from urlparse import urljoin
from azre.items import AzreItem
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request

import copy
import azre.settings
from scrapy import log
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher

class azre_scraper(Spider):

    name = "azre_scraper"

    base_url = 'http://services.azre.gov/'

    inputs_urls = 'azre_urls.txt'

    allowed_domains = ["services.azre.gov"]

    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    seen_ids = set()

    inputs_urls_lst = list()

    ds = None
    
    def __init__(self):
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        self.inputs_urls_lst = azre.settings.loadFileIntoLIst(self.inputs_urls)
        azre.settings.remaining_urls = copy.deepcopy(self.inputs_urls_lst)
        #self.ds = dblite.open(nzl_worldb_net_spider_item, 'sqlite://Murray_March_2015.sqlite:nzl_worlddb', autocommit=True)

    def spider_closed(spider, reason):
        log.msg('Reason to close the spider:' + reason)
        rem = azre.settings.remaining_urls
        f = open('azre_urls.txt', 'w')
        for r in rem:
            f.write(r + '\n')
        f.close()

    def start_requests(self):
        self.urls = MyLines4.loadFileIntoLIst(self.urls_file)
        self.urls = list(set(self.urls))
        fn2yc.settings.remaining_urls = self.urls

        #>>>>>>>>>>>>>>>>sqlite them using the id

        for url in self.urls:
            if url != '':

                listing_id = re.findall('/b/(\d*\-*i\d+)', url)
                if listing_id and len(listing_id) > 0:
                    listing_id = listing_id[0].strip()
                else:
                    listing_id = None

                items_found = self.ds.sql('SELECT * FROM nzl_worlddb WHERE "ListingID" = "%s"' % str(listing_id))
                items_found = items_found.gi_frame.f_locals['items']

                if items_found and len(items_found) > 0:
                    print 'Item found ' + listing_id
                    if url in fn2yc.settings.remaining_urls and len(url) > 1:
                        fn2yc.settings.remaining_urls.remove(url)
                        print str(len(fn2yc.settings.remaining_urls)) + ' Remaining'

                else:
                    source_url= url
                    meta = dict()
                    meta['source_url'] = source_url
                    yield Request(url=url.strip(),
                                  meta=copy.deepcopy(meta),
                                  headers=self.headers,
                                  callback=self.scrape_product_page)

    def parse_individual_page(self, response):
        sel = Selector(response)

        license_number = ''
        license_number_node = sel.xpath('//*[@id="DefaultContent_LabelLicenseNumber"]/text()')
        if license_number_node and len(license_number_node) > 0:
            license_number = license_number_node.extract()[0].strip()
            if not license_number in self.seen_ids:
                self.seen_ids.add(license_number)
            else:
                log.msg('ID scraped before!', log.INFO)
                yield

        name = ''
        name_node = sel.xpath('//*[@id="DefaultContent_LabelName"]/text()')
        if name_node and len(name_node) > 0:
            name = name_node.extract()[0].strip()

        nick_name = ''
        nick_name_node = sel.xpath('//*[@id="DefaultContent_LabelNickname"]/text()')
        if nick_name_node and len(nick_name_node) > 0:
            nick_name = nick_name_node.extract()[0].strip()

        license_status = ''
        license_status_node = sel.xpath('//*[@id="DefaultContent_LabelLicenseStatus"]/text()')
        if license_status_node and len(license_status_node) > 0:
            license_status = license_status_node.extract()[0].strip()

        license_type = ''
        license_type_node = sel.xpath('//*[@id="DefaultContent_LabelLicenseType"]/text()')
        if license_type_node and len(license_type_node) > 0:
            license_type = license_type_node.extract()[0].strip()

        pc_plc_name = ''
        pc_plc_name_node = sel.xpath('//*[@id="DefaultContent_LabelCorporationName"]/text()')
        if pc_plc_name_node and len(pc_plc_name_node) > 0:
            pc_plc_name = pc_plc_name_node.extract()[0].strip()

        original_date = ''
        original_date_node = sel.xpath('//*[@id="DefaultContent_LabelOriginalDate"]/text()')
        if original_date_node and len(original_date_node) > 0:
            original_date = original_date_node.extract()[0].strip()

        expiration_date = ''
        expiration_date_node = sel.xpath('//*[@id="DefaultContent_LabelExpirationDate"]/text()')
        if expiration_date_node and len(expiration_date_node) > 0:
            expiration_date = expiration_date_node.extract()[0].strip()

        item = AzreItem()

        item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
        item['Full_Name'] = name
        item['FirstName'] = ''
        item['LastName'] = ''
        if ',' in name:
            name_arr = name.split(',')
            item['FirstName'] = name_arr[1].strip()
            item['LastName'] = name_arr[0].strip()
        item['Nickname'] = nick_name
        item['License_Number'] = license_number
        item['License_Status'] = license_status
        item['License_Type'] = license_type
        item['PC_PLC_Name'] = pc_plc_name
        item['Original_Date'] = original_date
        item['Expiration_Date'] = expiration_date
        item['RefUrl'] = response.url

        employment_history = sel.xpath('//*[@id="ctl00_DefaultContent_RadGridEmployment_ctl00"]/tbody/tr')
        employment_history = sel.xpath('//*[@id="ctl00_DefaultContent_RadGridEmployment_ctl00"]/tbody/tr[1]')
        if employment_history and len(employment_history) > 0:
            for e in employment_history:
                cells = e.xpath('.//td')

                hist_url = cells[0].xpath('a/@href')
                if hist_url and len(hist_url) > 0:
                    hist_url = hist_url.extract()[0]
                    hist_url = urljoin('http://services.azre.gov/PublicDatabase/', hist_url)

                    cells = cells[1:]
                    company = cells[0].xpath('text()').extract()[0]
                    employment_type = cells[1].xpath('text()').extract()[0]

                    item['Employer'] = copy.deepcopy(company)
                    item['EmploymentType'] = copy.deepcopy(employment_type)

                    base_meta = dict()
                    base_meta['item'] = copy.deepcopy(item)
                    yield Request(url=hist_url,
                                  meta=copy.deepcopy(base_meta),
                                  headers=self.headers,
                                  callback=self.parse_entity_page)
        else:
            print response.url
            yield item

    def parse_entity_page(self, response):
        sel = Selector(response)
        item = response.meta['item']

        employer_mailing_address = ''
        employer_mailing_address_node = sel.xpath('//*[@id="DefaultContent_LabelMailingAddress"]/text()')
        if employer_mailing_address_node and len(employer_mailing_address_node) > 0:
            employer_mailing_address = employer_mailing_address_node.extract()[0]

        phone = ''
        phone_node = sel.xpath('//*[@id="DefaultContent_LabelPhone"]/text()')
        if phone_node and len(phone_node) > 0:
            phone = phone_node.extract()[0].strip()

        city = ''
        state = ''
        zip_code = ''
        employer_business_address = ''
        employer_business_address_node = sel.xpath('//*[@id="DefaultContent_LabelBusinessAddress"]/text()')
        if employer_business_address_node and len(employer_business_address_node) > 0:
            employer_business_address = employer_business_address_node.extract()[0]
            if employer_business_address_node and len(employer_business_address_node) > 1:
                employer_business_address2 = employer_business_address_node.extract()[1]
                if ',' in employer_business_address2:
                    bus_arr = employer_business_address2.split(',')
                    city = bus_arr[0].strip()
                    zip_code = re.findall('\d+', bus_arr[1])[0].strip()
                    state = bus_arr[1].replace(zip_code, '').strip()

        item['EmployerBusinessAddress'] = employer_business_address
        item['EmployerMailingAddress'] = employer_mailing_address
        item['City'] = city
        item['State'] = state
        item['ZipCode'] = zip_code
        item['Phone'] = phone

        yield item