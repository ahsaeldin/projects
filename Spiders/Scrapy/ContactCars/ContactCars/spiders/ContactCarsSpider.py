import re
import time
from urlparse import urljoin

from bs4 import BeautifulSoup
from scrapy.contrib.spiders import CrawlSpider, Rule
from scrapy.contrib.linkextractors.sgml import SgmlLinkExtractor
from ContactCars.items import ContactcarsItem


class IndeedSpider(CrawlSpider):

    name = "contactcars"
    allowed_domains = ["contactcars.com"]
    start_urls = ["http://www.contactcars.com/usedcars?direction=desc&pf=95000&pic=true&pt=105000&search=uc&sort=MakeYear"]

    rules = (
        #Rule(SgmlLinkExtractor(allow=('&page=2', ))),
        Rule(SgmlLinkExtractor(allow=('&page=\d+', )), callback='parse_item', follow=True),
        #Rule(SgmlLinkExtractor(allow=('start=\d+', )), callback='parse_item', follow=True),
    )

    def parse_item(self, response):
        #sel = Selector(response)
        #sites = sel.xpath('//ul/li')

        soup = BeautifulSoup(response.body)

        scrape_time = time.strftime('%Y-%m-%d %H:%M:%S')

        link = response.url

        rows = soup.find_all('div', {'class': 'panel custom_p_2'})

        for row in rows:
            make = ''
            make_tag = row.find('p', {'class': 'tit_2'})
            if make_tag:
                make = make_tag.text.strip()

            model_year_tags = row.find_all('p', {'class': re.compile('tit_3')})

            cc = ''
            year = ''
            model = ''
            for cell in model_year_tags:
                text = cell.text.strip()
                if re.search('2\d{3}', text) and not '2000' in text and not '2200' in text:
                    year = text
                elif re.search('\d{4}', text):
                    cc = text
                    cc = cc.replace('|', '').strip()
                else:
                    model = text.replace('-', '').strip()

            price_tag = row.find('span', {'class': 'orange_txt left'})
            if price_tag:
                price = price_tag.text

            if year:
                item = ContactcarsItem()
                item['Make'] = make.encode('utf-8')
                item['Year'] = year.encode('utf-8')
                item['Model'] = model.encode('utf-8')
                item['CC'] = cc.encode('utf-8')
                item['Price'] = price.encode('utf-8')
                item['ScrapeTime'] = scrape_time.encode('utf-8')
                item['URL'] = link.encode('utf-8')

                yield item
