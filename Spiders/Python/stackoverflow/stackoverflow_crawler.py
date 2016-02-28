#!/usr/bin/python

import re
import MyLines2
import requests
from bs4 import BeautifulSoup

def main():
    users_urls = []
    proxiesFile = 'proxies.dat'
    target_url = 'http://stackoverflow.com/users?tab=Reputation&filter=all'
    profiles_urls_file_name = 'stackoverflow_users.txt'
    headers = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
               "Accept-Encoding": "gzip,deflate,sdch",
               "Accept-Language": "en-US,en;q=0.8",
               "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.107 Safari/537.36"}

    proxies = MyLines2.loadProxies(proxiesFile)
    req = request_page(target_url, headers, proxies)

    if req.status_code == 200:
        pageSource = req.text
        if pageSource:
            extract_urls(pageSource)

            for i in xrange(2, 76875):
                next_url = 'http://stackoverflow.com/users?page=%s&tab=reputation&filter=all' % str(i)
                req = request_page(next_url, headers, proxies)
                if req.status_code == 200:
                    pageSource = req.text
                    if pageSource:
                        extract_urls(pageSource)
                        if i == 30:
                            break
                #print next_url

def request_page(target_url, headers, proxies):
    if proxies and len(proxies) > 0:
        rndProxy = 'http://' + MyLines2.PickRandomProxy(proxies)
        #print 'request using proxy::' + rndProxy
        req = requests.get(target_url, headers=headers, proxies={'http': rndProxy})
    else:
        req = requests.get(target_url, headers=headers)
    return req

def extract_urls(page_source):
    soup = BeautifulSoup(page_source)
    users_div = soup.find('div', id='user-browser')

    if users_div:
        users_anchors = users_div.find_all('a', href=re.compile('/users/.+?'))
        users_urls = ['http://stackoverflow.com' + x.attrs['href'] for x in users_anchors]
        users_urls = list(set(users_urls))
    for user in users_urls:
        print user

if __name__ == '__main__':
    main()
    print 'Done'