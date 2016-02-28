# -*- coding: utf-8 -*-

import re
import os
import time
import MyLines4
import HTMLParser
from bs4 import BeautifulSoup
from scrapy.spider import BaseSpider
from greendealapproved.items import GreendealapprovedItem

class GreendealapprovedSpider(BaseSpider):

    jobs_list = []
    name = "greendealapproved"
    urls_file = 'greendealapproved-urls.txt' #os.path.dirname(os.path.realpath(__file__)) + '/greendealapproved-urls.txt'
    allowed_domains = ["gdorb.decc.gov.uk"]
    start_urls = []

    def __init__(self):
        self.start_urls = MyLines4.loadFileIntoLIst(self.urls_file)

    def parse(self, response):
        return self.scrape_product_page(response.url, response.body)

    def scrape_product_page(self, url, body):
        soup = BeautifulSoup(body)

        company_name = ''
        company_name_tag = soup.find('h1')
        if company_name_tag:
            company_name = company_name_tag.text

        certification_id = ''
        certification_id_tag = soup.find('p', text=re.compile('Certification\sID:'))
        if certification_id_tag:
            certification_id = certification_id_tag.text
            certification_id = certification_id.replace('Certification ID:', '').strip()

        date_certified = ''
        date_certified_tag = soup.find('p', text=re.compile('Date\sCertified:'))
        if date_certified_tag:
            date_certified = date_certified_tag.text
            date_certified = date_certified.replace('Date Certified:', '').strip()

        telephone = ''
        telephone_tag = soup.find('p', text=re.compile('Telephone:'))
        if telephone_tag:
            telephone = telephone_tag.text
            telephone = telephone.replace('Telephone:', '').strip()

        website = ''
        websites_tags = soup.find_all('a', {'target': '_blank'})
        for wb in websites_tags:
            par = wb.find_parent()
            if par.name == 'p' and 'Website:' in par.text:
                website = wb.attrs['href'].strip()
                break

        email = ''
        email_tags = soup.find_all('script')
        for em in email_tags:
            par = em.find_parent()
            if par.name == 'p' and 'Email:' in par.text:
                inner = em.prettify()
                h = HTMLParser.HTMLParser()
                tamam = h.unescape(inner)
                tamam = tamam.replace("\r", ' ')
                tamam = tamam.replace("\n", ' ')
                em_tag = re.search('addy\d+.+?addy\d+\s=\saddy\d+.+?;', tamam)
                if em_tag:
                    email_line = em_tag.group(0)
                    email = re.sub("=|'|\s|addy\d+|;|\+", '', email_line)

        green_deal_services_offered = ''
        head2 = soup.find('h2', text='Green Deal services offered')
        if head2:
            green_deal_services_offered_tag = head2.find_next_sibling('ul')
            if green_deal_services_offered_tag:
                lis = green_deal_services_offered_tag.find_all('li')
                for li in lis:
                    if green_deal_services_offered == '':
                        green_deal_services_offered = li.text.strip()
                    else:
                        green_deal_services_offered += '|' + li.text.strip()

        name = ''
        name_matcher = re.search('Name:\s(\w+\s\w+)<', body)
        if name_matcher:
            name = name_matcher.group(1)

        address = ''
        address_head2 = soup.find('h2', text='Head Office')
        if address_head2:
            address_tag = address_head2.find_next_sibling('p')
            if address_tag:
                address = address_tag.text.strip()

        item = GreendealapprovedItem()
        item['DateExtracted'] = time.strftime('%Y-%m-%d %H:%M:%S')
        item['CompanyName'] = MyLines4.prepare_field(company_name)
        item['CertificationID'] = MyLines4.prepare_field(certification_id)
        item['DateCertified'] = MyLines4.prepare_field(date_certified)
        item['Telephone'] = MyLines4.prepare_field(telephone)
        item['Website'] = MyLines4.prepare_field(website)
        item['Email'] = MyLines4.prepare_field(email)
        item['Address'] = address
        item['Name'] = MyLines4.prepare_field(name)
        item['GreenDealServicesOffered'] = MyLines4.prepare_field(green_deal_services_offered)
        item['AdvertURL'] = MyLines4.prepare_field(url)

        return item
