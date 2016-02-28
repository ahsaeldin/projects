# -*- coding: utf-8 -*-

# Scrapy settings for food project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#

BOT_NAME = 'subscene_v3' #todo add bot name, don't forget to rename it in scrapy.cfg

SPIDER_MODULES = ['subscene_v3.spiders'] #todo
NEWSPIDER_MODULE = 'subscene_v3.spiders' #todo

# Crawl responsibly by identifying yourself (and your website) on the user-agent
#USER_AGENT = 'food (+http://www.yourdomain.com)'

CLOSESPIDER_ERRORCOUNT = 500 # close spider after 500 errors to avoid endless error log

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
      #  'subscene_v3.middlewares.ProxyMiddleware': 110,  #todo
        #'platt.middlewares.CustomUserAgentMiddleware': 100,
    }

ITEM_PIPELINES = { #todo
    'subscene_v3.pipelines.subscene_v3_pipleline': 300,
    #'myproject.pipelines.JsonWriterPipeline': 800,
}

LOG_LEVEL = 'INFO'

#LOG_FILE = 'log.txt'