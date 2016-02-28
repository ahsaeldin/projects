import base64
import random
from food.settings import PROXIES  #todo

class ProxyMiddleware(object):
    def process_request(self, request, spider):
        proxy = random.choice(PROXIES)
        if proxy.has_key('user_pass'):
            request.meta['proxy'] = "http://%s" % proxy['ip_port']
            encoded_user_pass = base64.encodestring(proxy['user_pass'])
            request.headers['Proxy-Authorization'] = 'Basic ' + encoded_user_pass
        else:
            proxy = "http://%s" % proxy['ip_port']
            request.meta['proxy'] = proxy
