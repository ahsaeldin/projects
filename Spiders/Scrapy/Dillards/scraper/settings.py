# -*- coding: utf-8 -*-

# Scrapy settings for scraper project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#

BOT_NAME = 'scraper'

SPIDER_MODULES = ['scraper.spiders']
NEWSPIDER_MODULE = 'scraper.spiders'

# Crawl responsibly by identifying yourself (and your website) on the user-agent
#USER_AGENT = 'scraper (+http://www.yourdomain.com)'

USER_AGENT_LIST = [
    'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.7 (KHTML, like Gecko) Chrome/16.0.912.36 Safari/535.7',
    'Mozilla/5.0 (Windows NT 6.2; Win64; x64; rv:16.0) Gecko/16.0 Firefox/16.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/534.55.3 (KHTML, like Gecko) Version/5.1.3 Safari/534.53.10'
]

HTTP_PROXY = 'http://127.0.0.1:8123'

# DOWNLOADER_MIDDLEWARES = {
#    #      'scraper.middlewares.RandomUserAgentMiddleware': 400,
#    #      'scraper.middlewares.ProxyMiddleware': 410,
#    #      'scrapy.contrib.downloadermiddleware.useragent.UserAgentMiddleware': None,
#     # Disable compression middleware, so the actual HTML pages are cached
# }

ITEM_PIPELINES = {
    'scraper.pipelines.FormatItemValuePipeline': 1,
    'scraper.pipelines.CsvExportPipeline': 2,
}

DOWNLOAD_DELAY = 0.25
CONCURRENT_ITEMS = 100
CONCURRENT_REQUESTS = 8
CONCURRENT_REQUESTS_PER_DOMAIN = 4
DOWNLOAD_TIMEOUT = 1000
REDIRECT_MAX_METAREFRESH_DELAY = 1000
REDIRECT_PRIORITY_ADJUST = -2  # keep this on, without it the scraper results will vary between 300 and 350 event_id's
                               # when actually the should be around 420