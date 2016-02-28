import re
import UnicodeRW
from selenium import webdriver

recipes = []

#keywords = ['salt','water','sugar','flour','bread','onion','pie','sandwich','pasta','salad']
keywords = ['pie','sandwich','pasta','salad']

driver = webdriver.Firefox()

def getLinks():
    isFound = False
    articleAnchors = driver.find_elements_by_xpath('//*[@id="search_results"]/article[*]/div/h2/a')
    for anchor in articleAnchors:
        url = anchor.get_attribute('href')
        recipes.append(url)
        isFound = True
    if isFound:
        return True
    else:
        return False


for keyword in keywords:
    print keyword
    url = 'http://www.seriouseats.com/search?site=recipes&term=%s' % keyword
    driver.get(url)
    getLinks()
    offset = 50
    while True:
        nextURL = 'http://www.seriouseats.com/search?term=%s&site=se&offset=%s' % (keyword, str(offset))
        offset = offset + 50
        driver.get(nextURL)
        if not getLinks():
            break
    f = open('links.txt', 'a')
    for recipe in recipes:
		if recipe:
			f.write(recipe + '\n')
    f.close()
    recipes = []

print 'Done'

