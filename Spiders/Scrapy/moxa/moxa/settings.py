# -*- coding: utf-8 -*-

# Scrapy settings for food project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#

BOT_NAME = 'moxa'

SPIDER_MODULES = ['moxa.spiders']
NEWSPIDER_MODULE = 'moxa.spiders'

# Crawl responsibly by identifying yourself (and your website) on the user-agent
#USER_AGENT = 'food (+http://www.yourdomain.com)'
USER_AGENT = 'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.81 Safari/537.36'
# BROWSERS = (
#     # Top most popular browsers in my access.log on 2009.02.12
#     # tail -50000 access.log |
#     #  awk -F\" '{B[$6]++} END { for (b in B) { print B[b] ": " b } }' |
#     #  sort -rn |
#     #  head -20
#     'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.6) Gecko/2009011913 Firefox/3.0.6',
#     'Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10.5; en-US; rv:1.9.0.6) Gecko/2009011912 Firefox/3.0.6',
#     'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.6) Gecko/2009011913 Firefox/3.0.6 (.NET CLR 3.5.30729)',
#     'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.6) Gecko/2009020911 Ubuntu/8.10 (intrepid) Firefox/3.0.6',
#     'Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US; rv:1.9.0.6) Gecko/2009011913 Firefox/3.0.6',
#     'Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US; rv:1.9.0.6) Gecko/2009011913 Firefox/3.0.6 (.NET CLR 3.5.30729)',
#     'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US) AppleWebKit/525.19 (KHTML, like Gecko) Chrome/1.0.154.48 Safari/525.19',
#     'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; .NET CLR 3.0.04506.30; .NET CLR 3.0.04506.648)',
#     'Mozilla/5.0 (X11; U; Linux x86_64; en-US; rv:1.9.0.6) Gecko/2009020911 Ubuntu/8.10 (intrepid) Firefox/3.0.6',
#     'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.5) Gecko/2008121621 Ubuntu/8.04 (hardy) Firefox/3.0.5',
#     'Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_5_6; en-us) AppleWebKit/525.27.1 (KHTML, like Gecko) Version/3.2.1 Safari/525.27.1',
#     'Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322)',
#     'Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 2.0.50727)',
#     'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)'
# )

CLOSESPIDER_ERRORCOUNT = 500 # close spider after 500 errors to avoid endless error log

def loadFileIntoLIst(fileName):
    with open(fileName) as f:
        _proxies = f.read().splitlines()
        _proxies = list(set(_proxies))
    return _proxies

PROXIES = list()
proxieslst = loadFileIntoLIst('proxies.txt')
for proxy in proxieslst:
    itm = {'ip_port': proxy}
    PROXIES.append(itm)

if len(PROXIES) > 0:
    DOWNLOADER_MIDDLEWARES = {
        'moxa.middlewares.ProxyMiddleware': 110,
        #'platt.middlewares.CustomUserAgentMiddleware': 100,
    }

ITEM_PIPELINES = {
    'moxa.pipelines.moxaPipeline': 300,
    #'myproject.pipelines.JsonWriterPipeline': 800,
}

LOG_LEVEL = 'INFO'

#LOG_FILE = 'log.txt'