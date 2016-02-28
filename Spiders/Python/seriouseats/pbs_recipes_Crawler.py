import re
import mechanize
from bs4 import BeautifulSoup

recipes = []
userAgent = 'Mozilla/5.0 (X11; Linux x86_64; rv:18.0) Gecko/20100101 Firefox/18.0 (compatible;)'
br = mechanize.Browser()

startURL = 'http://www.pbs.org/food/?post_type=food_recipes&s=+&Search=Search'

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
        return True
    else:
        return False

#Visit startURL
br.open(startURL)

lastLinkLst = list(br.links(text_regex=re.compile("Last")))

#Get total pages number
pagesNO = 0
if lastLinkLst and len(lastLinkLst) > 0:
    lastLink = lastLinkLst[0]
    pagesNoMatcher = re.search('/(\d{2,3})/', str(lastLink))
    if pagesNoMatcher:
        pagesNO = int(pagesNoMatcher.group(1))

#Visit them all in a loop
if pagesNO > 0:
    for num in range(1, pagesNO):
        url = 'http://www.pbs.org/food/page/%s/?post_type=food_recipes&s=+&Search=Search' % num
        #Collect URLs while looping
        print num
        br.open(url)
        getLinks(br)
        f = open('pbs-links.txt','a')
        for recipe in recipes:
            f.write(recipe+'\n')
        f.close()
        recipes = []