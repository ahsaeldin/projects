from bs4 import BeautifulSoup
from scrapy.http import Request, FormRequest
from scrapy.selector import HtmlXPathSelector
from scrapy.contrib.spiders.init import InitSpider

class rpyccspider(InitSpider):
    name = "rpycc"
    user_name = '2140'
    password = 'pink911'
    allowed_domains = ["rpycc.org"]
    login_page = 'http://www.rpycc.org/club/scripts/login/login.asp'
    start_urls = ['http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=A',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=B',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=C',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=D',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=F',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=G',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=H',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=I',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=J',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=K',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=L',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=M',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=N',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=O',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=P',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=Q',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=R',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=S',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=T',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=U',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=V',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=W',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=X',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=Y',
                  'http://www.rpycc.org/club/scripts/member/member_search.asp?GRP=23485&NS=MYLOCKER&LET=Z'
                 ]

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
        results_div = soup.find('div', {'class': 'memberdirectory-user-search-wrapper'})
        if results_div:
            all_profiles = results_div.find_all('a')
            for profile in all_profiles:
                url = profile.attrs['href']
                f = open('url.csv', 'ab')
                f.write(url + '\n')
                f.close()