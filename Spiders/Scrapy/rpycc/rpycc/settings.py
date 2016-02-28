# Scrapy settings for rpycc project
#
# For simplicity, this file contains only the most important settings by
# default. All the other settings are documented here:
#
#     http://doc.scrapy.org/en/latest/topics/settings.html
#

BOT_NAME = 'rpycc'

SPIDER_MODULES = ['rpycc.spiders']
NEWSPIDER_MODULE = 'rpycc.spiders'

# Crawl responsibly by identifying yourself (and your website) on the user-agent
#USER_AGENT = 'rpycc (+http://www.yourdomain.com)'

#DOWNLOAD_HANDLERS = {'http': 'scrapy.core.downloader.handlers.http.HttpDownloadHandler',
#                     'https': 'scrapy.core.downloader.handlers.http.HttpDownloadHandler',}