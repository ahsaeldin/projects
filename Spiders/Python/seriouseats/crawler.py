import re
import mechanize
from bs4 import BeautifulSoup

recipes = []

keywords = ['salt','water','sugar','flour','bread','onion','pie','sandwich','pasta','salad']

userAgent = 'Mozilla/5.0 (X11; Linux x86_64; rv:18.0) Gecko/20100101 Firefox/18.0 (compatible;)'

br = mechanize.Browser()

#br.addheaders = [('User-Agent', userAgent), ('Accept', '*/*')]

def getLinks(br2):
    isFound = False
    soup = BeautifulSoup(br2.response().read())
    articleAnchors = soup.find_all("a",attrs={"href":re.compile('.+?/recipes/.+?')})
    for anchor in articleAnchors:
        if anchor.text:
            url = anchor.attrs['href']
            if not url in recipes:
                recipes.append(url)
                isFound = True
    if isFound:
        return  True
    else:
        return False

for keyword in keywords:
    print keyword
    url = 'http://www.seriouseats.com/search?term=%s&site=recipes' % keyword
    br.open(url)
    getLinks(br)
    offset = 50
    while True:
        nextURL = 'http://www.seriouseats.com/search?term=%s&site=se&offset=%s' % (keyword, str(offset))
        offset = offset + 50
        br.open(nextURL)
        if not getLinks(br):
            break
    f = open('links.txt','a')
    for recipe in recipes:
        f.write(recipe+'\n')
    f.close()
    recipes = []