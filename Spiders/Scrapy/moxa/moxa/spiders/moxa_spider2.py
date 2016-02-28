# -*- coding: utf-8 -*-

import re
import time
import urllib
from urlparse import urljoin
from urlparse import urlparse
from moxa.items import MoxaItem
from scrapy.spider import Spider
from scrapy.selector import Selector
from os.path import splitext, basename
from scrapy.http.request import Request

import copy
import moxa.settings
from scrapy import log
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher

class azre_scraper(Spider):

    name = "moxa_spider2"

    base_url = 'http://www.moxa.com/'

    inputs_urls = 'part_numbers_and_urls.csv' #'moxa_urls.txt'  #'part_numbers.txt'  #

    allowed_domains = ['www.moxa.com']

    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    parts = dict()

    seen_ids = set()

    inputs_urls_lst = list()

    download_delay = 2

    def __init__(self):
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        if self.inputs_urls == 'part_numbers.txt':
            parts_lst = list()
            parts = moxa.settings.loadFileIntoLIst(self.inputs_urls)
            # for part in parts:
            #     part_no = part.replace('-', '_')
            #     part_no = part_no.replace(' ', '_').strip() + '.htm'
            #     parts_lst.append(part_no)
            # parts_lst = [urljoin('http://www.moxa.com/product/', x) for x in parts_lst]
            # for p in parts_lst:
            #     print p
            self.inputs_urls_lst = parts #copy.deepcopy(list(set(parts_lst)))
        elif self.inputs_urls == 'part_numbers_and_urls.csv':
            parts_and_urls = moxa.settings.loadFileIntoLIst(self.inputs_urls)
            for p in parts_and_urls:
                purl = p.split(',')
                if len(purl) == 2:
                    part_number = purl[0]
                    part_url = purl[1]
                    self.parts[part_number] = part_url
        else:
            self.inputs_urls_lst = moxa.settings.loadFileIntoLIst(self.inputs_urls)

        moxa.settings.remaining_urls = copy.deepcopy(self.inputs_urls_lst)

    def spider_closed(spider, reason):
        log.msg('Reason to close the spider:' + reason)
        rem = moxa.settings.remaining_urls
        f = open('moxa_remaining_urls.txt', 'w')
        for r in rem:
            f.write(r + '\n')
        f.close()

    #https://www.google.com.eg/webhp?hl=en&gws_rd=ssl#hl=en&q=85M-3801-T+site:www.moxa.com%2Fproduct

    def start_requests(self):
        # meta = dict()
        # meta['part_number'] = 'C168H'
        # yield Request(url='http://www.moxa.com/product/c168h.htm',
        #               meta=copy.deepcopy(meta),
        #               headers=self.headers,
        #               callback=self.parse_product_page)

        kokowowo = set()
        if self.inputs_urls == 'part_numbers_and_urls.csv':
            for k, v in self.parts.iteritems():
                meta = dict()
                meta['part_number'] = k
                # if k != 'C168H/PCI':
                #     continue
                if v != '' and v != 'phased out':
                    print v
                    kokowowo.add(v)
                    yield Request(url=v,
                                  dont_filter=True,
                                  meta=copy.deepcopy(meta),
                                  headers=self.headers,
                                  callback=self.parse_product_page,
                                  errback=lambda x: self.download_errback(x, v))
                else:
                    with open('remaining.txt', 'ab') as f2:
                        f2.write(k + '\n')
            print 'sdf'
        elif self.inputs_urls == 'part_numbers.txt':
            for part_no in self.inputs_urls_lst:
                query = 'q=%s site:www.moxa.com/product' % part_no
                query = query.replace(' ', '+')
                url = 'https://www.google.com/search?%s&num=1&ie=utf-8&oe=utf-8&aq=t&rls=org.mozilla:en-US:official&client=firefox-a&channel=fflb' % query
                print url
                meta = {}
                meta['part_number'] = part_no
                yield Request(url=url,
                              dont_filter=True,
                              meta=copy.deepcopy(meta),
                              headers=self.headers,
                              callback=self.parse_google_serp)
        elif self.inputs_urls == 'moxa_urls.txt':
            #here use every url you have for moxa and loop through them one by one so
            #check if it is a product page, then extract the part #

            #1.extract the part no from the url and validate it for the best guess.
            #2.go to the page and make sure it is a product page. GREATO!

            count = 0
            for url in self.inputs_urls_lst:
                if url != '':
                    xurl = url
                    if 'http://www.moxa.comhttp://www.moxa.com/' in xurl:
                        xurl = xurl.replace('http://www.moxa.comhttp://www.moxa.com/', 'http://www.moxa.com/')
                    yield Request(url=xurl,
                                  dont_filter=True,
                                  headers=self.headers,
                                  callback=self.parse_product_page)

                # if '-' in url or '_' in url:
                #     part_no = re.findall('/([\w\-_]+)\.htm', url)
                #     if part_no and len(part_no) > 0:
                #         part_no = part_no[0]
                #         count += 1
                #         print str(count)
                #         print part_no

    def parse_google_serp(self, response):
        ffff = open('gooogle as hehehehehe.csv', 'ab')
        sel = Selector(response)
        part_no = response.meta['part_number']
        results = sel.xpath('.//h3/a')
        is_found = False
        if results and len(results) > 0:
            results = results.extract()
            for res in results:
                url = re.findall('href="(.+?)"', res)
                text = re.findall('>([\w\-_\s]+)<', res)
                if url and len(url) > 0 and text and len(text) > 0:
                    url = url[0].strip()
                    text = text[0].strip()
                    if part_no in text:
                        is_found = True
                        ffff.write(part_no + ',' + url + '\n')
                        # yield Request(url=url,
                        #               headers=self.headers,
                        #               callback=self.parse_product_page)
        ffff.close()
        if not is_found:
            #print part_no
            #print text
            #print '-------------'
            with open('Failed_Part_no.txt', 'ab') as f_file:
                f_file.write(part_no + '\n')


    def parse_product_page(self, response):
        print response.url

        sel = Selector(response)

        sku = sel.xpath('.//h2/text()')
        if sku and len(sku) > 0:
            sku = sku.extract()[0].strip()

            sku2 = re.findall('/([\w\-_]+)\.htm', response.url)
            if sku2 and len(sku2):
                sku2 = sku2[0].strip()
                if '_' in sku2:
                    sku2 = sku2.replace('_', '-')
                    sku = sku2

            if self.inputs_urls == 'part_numbers_and_urls.csv':
                sku = response.meta['part_number']

            cats = list()
            bread = sel.xpath('//*[@id="header1_ment"]')
            if bread and len(bread) > 0:
                cats = bread.xpath('.//a/text()')
                if cats and len(cats) > 0:
                    cats = cats.extract()
                    cats = [x for x in cats if not x == 'Home' and not x == 'sign in' and not x == 'Products & Solutions']
                    last_cat = bread.xpath('.//div/div/text()')
                    if last_cat and len(last_cat) > 0:
                        last_cat = last_cat.extract()
                        last_cat = [re.sub('\r|\n', '', x) for x in last_cat]
                        last_cat = [x.replace('>', '').strip() for x in last_cat]
                        last_cat = list(set(last_cat))
                        last_cat = [x for x in last_cat if x != '']
                        if len(last_cat) == 1:
                            sku_as_last_cat = last_cat[0]
                            cats.append(sku_as_last_cat)

            img_filename = ''
            img = sel.xpath('//*[@id="picImg"]/@src')
            if img and len(img) > 0:
                img = img.extract()[0]
                img = urljoin(self.base_url, img)
                o = urlparse(img)
                filename, file_ext = splitext(basename(o.path))
                img_filename = sku + ' Moxa img' + '1'
                try:
                    if '/' in img_filename:
                        img_filename = img_filename.replace('/', '-')
                    urllib.urlretrieve(img, "images/" + img_filename + file_ext)
                except:
                    log.msg('Error in saving image ' + img_filename + ' possible invalid file name.')
                img_filename = img_filename + file_ext

            overview = ''
            features_and_benefits = ''
            features = sel.xpath('.//*[@id="tabCon"]')
            if features and len(features) > 0:
                headings4 = features.xpath('.//h4')
                if headings4 and len(headings4) > 0:
                    for h in headings4:
                        heading_text = h.xpath('text()')
                        if heading_text and len(heading_text) > 0:
                            heading_text = heading_text.extract()[0].strip()
                            if heading_text == 'Features and Benefits':
                                features_and_benefits = h.xpath('.//following-sibling::ul')
                                features_and_benefits = features_and_benefits.extract()
                                if features_and_benefits and len(features_and_benefits) > 0:
                                    features_and_benefits = features_and_benefits[0]
                                else:
                                    features_and_benefits = h.xpath('.//following-sibling::p')
                                    features_and_benefits = features_and_benefits.extract()
                                    if features_and_benefits and len(features_and_benefits) > 0:
                                        features_and_benefits = features_and_benefits[0]
                                    else:
                                        features_and_benefits = ''
                                if features_and_benefits != '':
                                    features_and_benefits = '<h4>Features and Benefits</h4>\n' + features_and_benefits
                            elif heading_text == 'Overview':
                                overview = h.xpath('.//following-sibling::p')
                                overview = overview.extract()
                                overview = '\r\n'.join(overview)
                                overview = '<h4>Overview</h4>' + '\r\n' + overview

            models_str_list = list()
            models = sel.xpath('//*[@id="ctl14_divAv"]')
            if models and len(models) > 0:
                models = models.xpath('.//table')
                if models and len(models) > 0:
                    models = models.extract()[0].split('\r\n')
                    models = [x.strip() for x in models if not '<th style="width:30px;">' in x and not "checkbox" in x]
                    models = [x for x in models if not x == '']
                    for m in models:
                        if 'class="w105 col009a82">' in m:
                            i_sku = re.findall('>(.+?)</td>', m)
                            if i_sku and len(i_sku) > 0:
                                i_sku = i_sku[0]
                                if '/' in i_sku:
                                    i_sku = i_sku.replace('/', '-')
                                models_str_list.append('<td valign="middle" class="w105 col009a82"><a href="http://www.piecel.eu/%s/">%s</a></td>' % (i_sku, i_sku))
                        else:
                            models_str_list.append(m)
                    models = '\n'.join(models_str_list)
                    #models = models.replace('cellspacing="5">\n</tr>', 'cellspacing="5">')
                    models = '<h3 class="firstTit">Available Models</h3>' + '\r\n' + models
                    if not '<tr' in models:
                        models = ''

            else:
                models = ''

            comp_models_str_list = list()
            comp_models = sel.xpath('//*[@id="ctl14_divCompatible"]')
            if comp_models and len(comp_models) > 0:
                comp_models = comp_models.xpath('.//table')
                if comp_models and len(comp_models) > 0:
                    comp_models = comp_models.extract()[0]
                    comp_models = comp_models.split('\r\n')
                    comp_models = [x.strip() for x in comp_models if not '<th style="width:30px;">' in x and not "checkbox" in x]
                    comp_models = [x for x in comp_models if not x == '']
                    for m in comp_models:
                        if 'class="w105 col009a82">' in m:
                            c_sku = re.findall('>(.+?)</td>', m)
                            if c_sku and len(c_sku) > 0:
                                c_sku = c_sku[0]
                                if '/' in c_sku:
                                    c_sku = c_sku.replace('/', '-')
                            comp_models_str_list.append('<td valign="middle" class="w105 col009a82"><a href="http://www.piecel.eu/%s/">%s</a></td>' % (c_sku.strip(),
                                                                                                                                                       c_sku.strip()))
                        else:
                            comp_models_str_list.append(m)
                    comp_models = '\n'.join(comp_models_str_list)
                    #models = models.replace('cellspacing="5">\n</tr>', 'cellspacing="5">')
                    comp_models = '<h3 class="firstTit">Compatible Models</h3>' + '\r\n' + comp_models
                    if not '<tr' in comp_models:
                        comp_models = ''
            else:
                comp_models = ''

            if len(comp_models) == 0:
                comp_models = ''

            weight = ''
            spec_str = ''
            dimensions = ''
            if features and len(features) > 0:
                specs = features.xpath('.//table')
                if specs and len(specs) > 0:
                    for spec in specs:
                        table = spec.extract()
                        #if 'Hardware Specifications' in table or 'Hardware' in table or 'Antenna Characteristics' in table:
                        if '<table' in table and 'width: 628px' in table and 'font-size: 12px' in table:
                            spec_str = table
                            specs_table_rows = spec.xpath('.//tr')
                            for r in specs_table_rows:
                                cells = r.xpath('.//td/text()')
                                if cells and len(cells) > 0:
                                    cells = cells.extract()
                                    if len(cells) > 1:
                                        if cells[0] == 'Weight':
                                            weight = cells[1]
                                        elif cells[0] == 'Dimensions':
                                            dimensions = cells[1]
                            break

            related_products_str = ''
            related_products = sel.xpath('//*[@id="rightMenuBottom1_divRelated"]')
            if related_products and len(related_products) > 0:
                related_products = related_products.xpath('.//div[@class="ow_rightCon"]')
                if related_products and len(related_products) > 0:
                    related_products = related_products.xpath('.//a/text()')
                    if related_products and len(related_products) > 0:
                        related_products = related_products.extract()
                        related_products = [re.sub('\r|\n', '', x).strip() for x in related_products]
                        related_products_str = ','.join(related_products)

            item = MoxaItem()
            item['SKU'] = sku
            item['MainCategory'] = cats[0] if len(cats) > 0 else ''
            item['Subcategory1'] = cats[1] if len(cats) > 1 else ''
            item['Subcategory2'] = cats[2] if len(cats) > 2 else ''
            item['Subcategory3'] = cats[3] if len(cats) > 3 else ''
            item['Subcategory4'] = cats[4] if len(cats) > 4 else ''
            item['Images'] = img_filename
            item['FeaturesAndBenefits'] = features_and_benefits.replace('"', "'")
            item['Overview'] = overview.replace('"', "'")
            item['AvailableModels'] = models.replace('"', "'")
            item['CompatibleModules'] = comp_models.replace('"', "'")
            item['HardwareSpecifications'] = spec_str.replace('"', "'")
            item['Weight'] = weight
            item['Dimensions'] = dimensions
            item['RelatedProducts'] = related_products_str

            data_sheet = sel.xpath('//*[@id="ctl14_pDatasheet"]/a/@href')
            if data_sheet and len(data_sheet) > 0:
                data_sheet = data_sheet.extract()[0]
                data_sheet = urljoin(self.base_url, data_sheet)
                meta = dict()
                meta['item'] = item
                yield Request(url=data_sheet,
                              meta=copy.deepcopy(meta),
                              headers=self.headers,
                              dont_filter=True,
                              callback=self.parse_datasheet_page,
                              errback=lambda x: self.download_errback(x, data_sheet))


    def parse_datasheet_page(self, response):
        sel = Selector(response)
        item = response.meta['item']
        sku = item['SKU']

        datasheet_filename = ''
        datasheet = sel.xpath('.//p/a')
        if datasheet and len(datasheet) > 0:
            for d in datasheet:
                link_href = d.xpath('@href')
                if link_href and len(link_href) > 0:
                    link_href = link_href.extract()[0]
                    if 'DownloadFile.aspx?type' in link_href:
                        link_href = urljoin('http://www.moxa.com/support/', link_href)
                        o = urlparse(link_href)
                        datasheet_filename = sku + ' Moxa datasheet' + '.pdf'
                        datasheet_filename = datasheet_filename.replace('/', '-')
                        urllib.urlretrieve(link_href, "datasheets/" + datasheet_filename)

        item['Datasheet'] = datasheet_filename

        yield item

    def download_errback(self, e, url):
        print url