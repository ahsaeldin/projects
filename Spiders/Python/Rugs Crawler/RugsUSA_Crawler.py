import re
import MyLines1
import requests

proxiesFile = 'proxies.dat'
catsUrlsFileName ='URLs.dat'
targetURL = 'http://api.searchspring.net/api/search/search.json?callback=jQuery17109579167992342263_1396464677850&method=search&q=&format=json&resultLayout=grid&resultsPerPage=50&domain=http%3A%2F%2Fwww.rugsusa.com%2Frugsusa%2Fcontrol%2Fsearch-rugs%23%2F%3F_%3D1%26filter.brand%3DCapel%26filter.brand%3DDynamic%2520Rugs%26filter.brand%3DColonial%2520Mills%26filter.brand%3DHomespice%2520Decor%26filter.brand%3DKas%2520Oriental%26filter.brand%3DMomeni%26filter.brand%3DOriental%2520Weavers%26filter.brand%3DLuxor%26filter.brand%3DTayse%2520Rugs%26filter.brand%3DThe%2520Rug%2520Market%26page%3D1&referer=&userId=F7110DA8-0492-4D31-A648-2A64D03B6EB4&facebook=0&websiteKey=d287e051460dfcc681456e26b5f67080&_=1396464678459&filter.brand=Capel&filter.brand=Dynamic+Rugs&filter.brand=Colonial+Mills&filter.brand=Homespice+Decor&filter.brand=Kas+Oriental&filter.brand=Momeni&filter.brand=Oriental+Weavers&filter.brand=Luxor&filter.brand=Tayse+Rugs&filter.brand=The+Rug+Market&page=koko&bgfilter.product_type=Rug&intellisuggest=1&requestId=74897105828858910&requestCount=0'
headers = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
           "Accept-Encoding": "gzip,deflate,sdch",
           "Accept-Language": "en-US,en;q=0.8",
           "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.107 Safari/537.36"}

prodURLs = []
for i in xrange(1, 119):
    NextURL = targetURL.replace('page=koko', 'page=' + str(i))
    req = requests.get(NextURL)
    if req.status_code == 200:
        pageSource = req.text
        pageSource = re.sub('jQuery\d+_\d+\(', '', pageSource)
        pageSource = re.sub('\);', '', pageSource)
        links = re.findall('<a\shref=."(.+?)"', pageSource)
        for link in links:
            _link = link.replace("\/", "/")
            _link = _link.replace("\\", "")
            prodURLs.append(_link)
            print _link

prodURLs = set(prodURLs)
MyLines1.SaveLstIntoFile(prodURLs, catsUrlsFileName)
print 'Done'