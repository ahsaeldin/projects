# -*- coding: utf-8 -*-

# Scrapy settings for food project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#

BOT_NAME = 'aido'

SPIDER_MODULES = ['aido.spiders']
NEWSPIDER_MODULE = 'aido.spiders'

# Crawl responsibly by identifying yourself (and your website) on the user-agent
#USER_AGENT = 'food (+http://www.yourdomain.com)'


def loadFileIntoLIst(fileName):
    with open(fileName) as f:
        _proxies = f.read().splitlines()
    return _proxies

PROXIES = list()
proxieslst = loadFileIntoLIst('proxies.txt')
for proxy in proxieslst:
    itm = {'ip_port': proxy}
    PROXIES.append(itm)

if len(PROXIES) > 0:
    DOWNLOADER_MIDDLEWARES = {
        #'aido.middlewares.ProxyMiddleware': 110,
        #'platt.middlewares.CustomUserAgentMiddleware': 100,
    }