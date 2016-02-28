import re
from items import RpyccItem
from bs4 import BeautifulSoup
from scrapy.http import Request, FormRequest
from scrapy.selector import HtmlXPathSelector
from scrapy.contrib.spiders.init import InitSpider

class rpyccspider(InitSpider):
    name = "rypceescraper"
    user_name = '2140'
    password = 'pink911'
    allowed_domains = ["rpycc.org"]
    login_page = 'http://www.rpycc.org/club/scripts/login/login.asp'
    start_urls = []

    def __init__(self):
        self.start_urls = self.load_start_urls('rpycc profiles urls')

    def init_request(self):
        #"""This function is called before crawling starts."""
        return Request(url=self.login_page, callback=self.login)

    def login(self, response):
        #"""Generate a login request."""
        return FormRequest.from_response(response,
                                         formname='frmLogin',
                                         formdata={'user': self.user_name, 'pw': self.password},
                                         callback=self.check_login_response)

    def check_login_response(self, response):
        #"""Check the response returned by a login request to see if we are successfully logged in."""

        if "Logout" in response.body:
            self.log("\n\n\nSuccessfully logged in. Let's start crawling!\n\n\n")
            # Now the crawling can begin..

            return self.initialized() # ****THIS LINE FIXED THE LAST PROBLEM*****

        else:
            self.log("\n\n\nFailed, Bad times :(\n\n\n")
            # Something went wrong, we couldn't log in, so nothing happens.

    def parse(self, response):
        print response.url
        soup = BeautifulSoup(response.body)

        profile_url = response.url

        #<h1 class="memberprofile-user-name">MR GREGORY TALBOTT</h1>
        name = ''
        name_tag = soup.find('h1', {'class': 'memberprofile-user-name'})
        if name_tag:
            name = self.prepare_field(name_tag.text)

        #<span class="memberprofile-user-location">BOCA RATON, FL</span>
        location = ''
        location_tag = soup.find('span', {'class': 'memberprofile-user-location'})
        if location_tag:
            location = self.prepare_field(location_tag.text)

        #<span class="memberprofile-user-email"><a href="mailto:gtalbott@talbottrealty.com">gtalbott@talbottrealty.com</a></span>
        email = ''
        email_tag = soup.find('span', {'class': 'memberprofile-user-email'})
        if email_tag:
            email = self.prepare_field(email_tag.text)

        #<span class="memberprofile-user-content-misc"><em>Member Since:</em> 2/27/2003</span>
        member_since = ''
        member_since_tag = soup.find('span', {'class': 'memberprofile-user-content-misc'})
        if member_since_tag:
            member_since = self.prepare_field(member_since_tag.text)

        #<address>250 NE 5TH AVENUE<br><span class="memberprofile-primary-adr-city-state">BOCA RATON, FL</span>  33432<br><br></address>
        address = ''
        address_tag = soup.find('address')
        if address_tag:
            address = self.prepare_field(address_tag.text)

        #<span class="memberprofile-user-address-phone"><em>Home:&nbsp;</em>561-289-7157 C</span>
        address_phone1 = ''
        address_phone2 = ''
        address_phone3 = ''
        address_phone1_tags = soup.find_all('span', {'class': 'memberprofile-user-address-phone'})
        count = 0
        for addr in address_phone1_tags:
            count += 1
            if count == 1:
                address_phone1 = self.prepare_field(addr.text)
            elif count == 2:
                address_phone2 = self.prepare_field(addr.text)
            elif count == 3:
                address_phone3 = self.prepare_field(addr.text)

        #<div class="memberprofile-user-pic-mask"><a href="/Images/Library/1792.jpg" title="Click to Enlarge this Photo of: MR DAVID ADAMS"></a></div>
        profile_image = ''
        profile_image_tag = soup.find('div', {'class': 'memberprofile-user-pic-mask'})
        if profile_image_tag:
            href_tag = profile_image_tag.find('a')
            if href_tag:
                href = 'http://www.rpycc.org' + href_tag.attrs['href']
                profile_image = self.prepare_field(href)

        item = RpyccItem()
        item['name'] = name
        item['location'] = location
        item['email'] = email
        item['member_since'] = member_since
        item['address'] = address
        item['address_phone1'] = address_phone1
        item['address_phone2'] = address_phone2
        item['address_phone3'] = address_phone3
        item['profile_image'] = profile_image
        item['profile_url'] = profile_url

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