# -*- coding: utf-8 -*-

# Scrapy settings for food project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#

BOT_NAME = 'leafin'

SPIDER_MODULES = ['leads_finder.spiders']
NEWSPIDER_MODULE = 'leads_finder.spiders'

# Crawl responsibly by identifying yourself (and your website) on the user-agent

#USER_AGENT = 'food (+http://www.yourdomain.com)'

def loadFileIntoLIst(fileName):
    with open(fileName) as f:
        _proxies = f.read().splitlines()
    return _proxies

PROXIES = list()

database_session = None

proxieslst = loadFileIntoLIst('proxies.txt')
for proxy in proxieslst:
    itm = {'ip_port': proxy}
    PROXIES.append(itm)

print 'Proxies only if found on the file.<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<'
if len(PROXIES) > 0:
    DOWNLOADER_MIDDLEWARES = {
        'leads_finder.middlewares.ProxyMiddleware': 110,
    }
# else:
#     print 'WARNING:No proxies found!!'
#     print 'Using Crawlera<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<'
#     print 'w el3adm l omak'
#     DOWNLOADER_MIDDLEWARES = {
#         #'leadgen.middlewares.ProxyMiddleware': 110,
#         'scrapy_crawlera.CrawleraMiddleware': 200,
#         #'platt.middlewares.CustomUserAgentMiddleware': 100,
#     }
#
#     CRAWLERA_ENABLED = True
#     CRAWLERA_USER = 'dec5233df8bc4dd7aa1223140dd06d6a'
#     CRAWLERA_PASS = ''

ITEM_PIPELINES = {
    'leads_finder.pipelines.LeadFinSqlAlchemyPipeline': 400,
}

LOG_LEVEL = 'INFO'

#LOG_FILE = 'log.txt'
