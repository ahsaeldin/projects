# -*- coding: utf-8 -*-

import re
import time
import copy
import datetime
from scrapy.spider import Spider
from scrapy.selector import Selector
from scrapy.http.request import Request
from leads_finder.items import LeadItem

class SingleLeadPageScraper(Spider):

    # Use this one if you already have the urls of the leads page direclty (asins, ids etc...)
    name = "SingleLeadPageScraper"
    keywords = 'input_codes.txt'
    allowed_domains = ['amazon.com']
    base_url = 'http://www.amazon.com/'
    headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Encoding': 'gzip, deflate',
               'Accept-Language': 'en-US,en;q=0.5',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:32.0) Gecko/20100101 Firefox/32.0'}

    def log_message(self, message):
        f = open('Missing.txt', 'ab')
        f.write(message + '\n')
        f.close()

    def start_requests(self):
        file = open(self.keywords)
        codes_list = file.read().splitlines()
        codes_list = list(set(codes_list))
        file.close()

        codes_count = len(codes_list)

        print'Starting...'

        meta = dict()
        codes_counter = 0
        for code in codes_list:
            if code == '' or len(code.strip()) != 10:
                continue
            codes_counter += 1
            url = 'http://www.amazon.com/dp/%s/' % code.strip()

            meta['code'] = code
            print ('Code:' + str(codes_counter) + '/' + str(codes_count) +
                   ' -- Extracted >>|' + '<< AMA YUHO POHO TOHO - AMA')
            yield Request(url=url,
                          meta=copy.deepcopy(meta),
                          headers=self.headers,
                          callback=self.parse_overview_page)

    def parse_overview_page(self, response):

        #@@>>> Before the above "yield Request("...
            # 1. Check the database if the ASIN is already there and skip it.
            # 2. Sniff the entire page for more ASINs man. and add them to a new file.

        # 1. Check missing fields below
        #&And&
        # 2. Valid int fields [Amazon_Best_Sellers_Rank, publication_year, pages_length, ISBN_10, ISBN_13]

        sel = Selector(response)

        lang = ''
        edition = ''
        isbn_13 = ''
        isbn_10 = ''
        pub_year = ''
        cover_type = ''
        pub_date = ''
        publisher = ''
        paper_count = ''
        amazon_rank = ''
        shipping_weight = ''

        is_kindle = False

        #Kindle Edition
        book_type = sel.xpath('//*[@id="booksTitle"]')
        if book_type and len(book_type) > 0:
            book_type = book_type.extract()[0]
            if 'Kindle Edition' in book_type:
                is_kindle = True

        title_tag = sel.xpath('//*[@id="productTitle"]/text()')
        if title_tag and len(title_tag) > 0:
            title = title_tag.extract()[0].strip()
        else:
            title = ''

        details_found = False
        book_details = sel.xpath('//*[@id="productDetailsTable"]')
        if book_details and len(book_details) > 0:
            details_found = True
            details = book_details.xpath('.//li')
            for detail in details:
                detail_name = detail.xpath('b/text()')
                if detail_name and len(detail_name) > 0:
                    detail_name = detail_name.extract()[0].strip()

                    detail_value = detail.xpath('text()')
                    if detail_value and len(detail_value) > 0:

                        detail_value = detail_value.extract()[0].strip()

                        check_pages = re.findall('[\d,]+ pages', detail_value)
                        if check_pages and len(check_pages) > 0:
                            paper_count = check_pages[0].replace('pages', '')
                            cover_type = detail.xpath('.//b/text()')
                            if cover_type and len(cover_type) > 0:
                                cover_type = cover_type.extract()[0].replace(':', '')
                            else:
                                cover_type = ''

                        if 'ISBN-10' in detail_name:
                            isbn_10 = detail_value
                        elif 'Amazon Best Sellers Rank' in detail_name:
                            amazon_rank = detail.extract()
                            amazon_rank = re.findall('#[\d,]+ in Books', amazon_rank)
                            if amazon_rank and len(amazon_rank) > 0:
                                amazon_rank = amazon_rank[0].replace('in Books', '').replace(' ', '').replace('#', '').replace(',', '')
                            else:
                                amazon_rank = ''
                        elif 'ISBN-13' in detail_name:
                            isbn_13 = detail_value.replace('-', '').strip()
                        elif 'Language' in detail_name:
                            lang = detail_value.strip()
                        elif 'Shipping Weight' in detail_name:
                            shipping_weight = detail_value.replace('(', '').strip()
                        elif 'Paperback' in detail_name:
                            paper_count = detail_value.replace('pages', '').strip()
                            cover_type = 'Paperback'
                        elif 'Hardcover' in detail_name:
                            paper_count = detail_value.replace('pages', '').strip()
                            cover_type = 'Hardcover'
                        elif 'Publisher' in detail_name:
                            publisher = detail_value.replace('(', '').strip()
                            pub_date = re.findall('[a-zA-Z]+\s\d{1,2},\s\d{4}', publisher)
                            if pub_date and len(pub_date) > 0:
                                pub_date = pub_date[0].strip()
                                publisher = publisher.replace(pub_date, '').replace('(', '').replace(')', '').strip()
                                edition = re.findall(';[\d+a-zA-Z\s]+edition', publisher)
                                if edition and len(edition) > 0:
                                    edition = edition[0].replace(';', '').strip()
                                else:
                                    edition = ''
                                pub_year = re.findall('\d{4}', pub_date)
                                if pub_year and len(pub_year) > 0:
                                    pub_year = pub_year[0].strip()
                                pub_date = datetime.datetime.strptime(pub_date, '%B %d, %Y').strftime('%m/%d/%Y')
                            else:
                                pub_date = ''

        if amazon_rank == '':
            amazon_rank = sel.xpath('//*[@id="SalesRank"]')
            if amazon_rank and len(amazon_rank) > 0:
                amazon_rank = amazon_rank.extract()[0]
                amazon_rank = re.findall('#[\d,]+ in Books', amazon_rank)
                if amazon_rank and len(amazon_rank) > 0:
                    amazon_rank = amazon_rank[0].replace('in Books', '').replace(' ', '').replace('#', '').replace(',', '')
                else:
                    amazon_rank = ''
            else:
                amazon_rank = ''

        byline = sel.xpath('//*[@id="byline"]')
        if byline and len(byline) > 0:
            byline = byline.xpath('.//span/a/text()')
            if byline and len(byline) > 0:
                byline = byline.extract()
                byline = [x.strip() for x in byline]
                byline = '|'.join(byline)
        else:
            byline = ''

        skip = False
        print response.url
        if not details_found:
            self.log_message('not details found')
            print'\tnot details found'
            skip = True

        if amazon_rank == '':
            self.log_message('Cannot find amazon_rank')
            print'\tCannot find amazon_rank.'
            amazon_rank = -1
            skip = True
        else:
            amazon_rank = int(amazon_rank)

        if title == '':
            self.log_message('Cannot find title')
            print'\tCannot find title.'
            skip = True

        if byline == '':
            self.log_message('Cannot find byline')
            print'\tCannot find byline.'
            skip = True

        if isbn_10 == '':
            self.log_message('Cannot find isbn_10')
            print'\tCannot find isbn_10.'
            skip = True

        if isbn_13 == '':
            self.log_message('Cannot find isbn_13.')
            print'\tCannot find isbn_13.'
            skip = True

        if lang == '':
            self.log_message('Cannot find lang')
            print'\tCannot find lang.'
            skip = True

        if shipping_weight == 'Cannot find shipping_weight':
            print'\tCannot find shipping_weight.'
            skip = True

        if publisher == '':
            self.log_message('Cannot find publisher')
            print'\tCannot find publisher.'
            skip = True

        if pub_date == '':
            self.log_message('Cannot find pub_date')
            print'\t\tCannot find pub_date.'
            skip = True

        if pub_year == '':
            self.log_message('Cannot find pub_year')
            print'\t\tCannot find pub_year.'
            skip = True

        if paper_count == '':
            self.log_message('Cannot find paper_count')
            print'\tCannot find paper_count.'
            skip = True

        if edition == '':
            self.log_message('Edition not found')
            print'\tEdition not found'
            skip = True

        if cover_type == '':
            book_type = sel.xpath('//*[@id="title"]/span[2]/text()')
            if book_type and len(book_type) > 0:
                cover_type = book_type.extract()[0].strip()
            else:
                self.log_message('Cannot find cover type')
                print'\tCannot find cover type'
                skip = True

        is_excluded = 0
        exclude_reason = ''
        if is_kindle or cover_type == 'Audio CD' or amazon_rank > 6000000 or amazon_rank == -1:
            is_excluded = 1
            if is_kindle:
                exclude_reason = 'Kindle Book'
            elif cover_type == 'Audio CD':
                exclude_reason = 'Audio CD'
            elif amazon_rank > 6000000:
                exclude_reason = 'amazon_rank > 6,000,000'
            elif amazon_rank == -1:
                exclude_reason = 'Best Seller Rank Not Found'

        item = LeadItem()
        item['Edition'] = edition
        item['cover_type'] = cover_type
        item['TimeStamp'] = time.strftime('%Y-%m-%d %H:%M:%S')
        item['Book_ID'] = response.meta['code']
        item['Amazon_Best_Sellers_Rank'] = amazon_rank
        item['ISBN_10'] = isbn_10
        item['ISBN_13'] = isbn_13
        item['ASIN'] = response.meta['code']
        item['Amazon_title'] = title
        item['pages_length'] = paper_count
        item['Amazon_author'] = byline
        item['Amazon_publisher'] = publisher
        item['publication_date'] = pub_date
        item['publication_year'] = pub_year
        item['shipping_weight'] = shipping_weight
        item['Amazon_language'] = lang
        item['is_excluded'] = is_excluded
        item['exclude_reason'] = exclude_reason

        if not skip and is_excluded == 0:
            yield item



