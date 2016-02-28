# -*- coding: utf-8 -*-

import re
from urlparse import urljoin
from bs4 import BeautifulSoup
from scrapy.spider import BaseSpider
from resortdata.items import ResortdataItem

class ResortDataSpider(BaseSpider):

    start_urls = []
    name = "resortdata"
    urls_file = '/home/ahmed/Desktop/urls.csv'
    allowed_domains = ["resortdata.com"]

    def __init__(self):
        self.start_urls = self.load_start_urls(self.urls_file)

    def parse(self, response):
        soup = BeautifulSoup(response.body)

        if soup:
            all_tables = soup.find_all('table')
            for table in all_tables:
                h1_tag = table.find('h1')
                if h1_tag:
                    h1 = h1_tag.text

                    content_table = h1_tag.findNextSibling('table')
                    if content_table:
                        body = content_table.text

                    banner_image_url = ''
                    imgs = table.find_all('img')
                    for img in imgs:
                        src = img.attrs['src']

                        if 'FacebookWeb.jpg' in src or 'TwitterWeb.jpg' in src:
                            continue

                        h = img.attrs['height']
                        w = img.attrs['width']

                        if h > 200 and w >=950 and not 'index_11' in src:
                           banner_image_url = urljoin(response.url, src)
                           break

                    item = ResortdataItem()
                    item['h1'] = self.prepare_field(h1)
                    item['body_stripped'] = self.prepare_field(body)
                    item['banner_image_url'] = self.prepare_field(banner_image_url)
                    item['url'] = response.url

                    return item

    def prepare_field(self, field):
        #field = field.encode('utf-8')
        field = re.sub('\n|\r|\t', ' ', field).strip()
        return field

    def load_start_urls(self, filename):
        urls = []
        with open(filename) as f:
            content = f.readlines()
            for row in content:
                urls.append(row.strip())
        return urls
