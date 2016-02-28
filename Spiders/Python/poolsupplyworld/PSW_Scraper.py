#python_template

import re
import csv
import requests
from bs4 import BeautifulSoup

targetURL = 'http://www.poolsupplyworld.com/manufacturers.htm'

#1.Get All Manufacturers URLs
manufacturersURLs = []
req = requests.get(targetURL)
if req.status_code == 200:
    pageSource = req.text
    if pageSource:
        soup = BeautifulSoup(pageSource)
        manufacturersTags = soup.findChildren('div', {'class': 'manufacturersListCol'})
        for manufacturersTag in manufacturersTags:
            manufacturersAnchors = manufacturersTag.findAll('a')
            for manufacturersAnchor in manufacturersAnchors:
                href = manufacturersAnchor.attrs['href']
                if href:
                    if not href in manufacturersURLs:
                        manufacturersURLs.append(href)
                        print 'manufacturer URL::' + href

#2.Loop through each manufacturer, and get its inner categories.
manufacturersCats = []
for manufacturersURL in manufacturersURLs:
    req = requests.get(manufacturersURL)
    if req.status_code == 200:
        pageSource = req.text
        if pageSource:
            soup = BeautifulSoup(pageSource)
            catsTag = soup.find('div', {'id': 'subcats'})
            if catsTag:
                catsAnchors = catsTag.findAll('a')
                for catsAnchor in catsAnchors:
                    href = catsAnchor.attrs['href']
                    if href:
                        if not 'http://www.poolsupplyworld.com' in href:
                            href = 'http://www.poolsupplyworld.com' + href
                        if not href in manufacturersCats:
                            manufacturersCats.append(href)
                            print 'manufacturer Cat::' + href
        #break#######################################################>>>>>>>>>>>>>> Remove this one.

#3.loop through Cats and get the Prods URLs *AJAX*
prodURLs = []
for manufacturersCat in manufacturersCats:
    req = requests.get(manufacturersCat)
    if req.status_code == 200:
        pageSource = req.text
        if pageSource:
            soup = BeautifulSoup(pageSource)
            productsTag = soup.find('div', {'id': 'product_display'})
            if productsTag:
                productsAnchors = productsTag.findAll('a')
                for productsAnchor in productsAnchors:
                    href = productsAnchor.attrs['href']
                    if href:
                        if not 'http://www.poolsupplyworld.com' in href:
                            href = 'http://www.poolsupplyworld.com' + href
                        if not href in prodURLs:
                            prodURLs.append(href)
                            print 'prod URL::' + href

            #Get cat id in order to format the AJAX call.
            catId = re.search('category_id=(\d+)', pageSource)
            if catId:
                counter = 2
                looping = True
                while looping:
                    moreProdsURL = 'http://www.poolsupplyworld.com/ajax/ajaxdriver.cfm?d=search&cont=search_v4&json=true&category_id=%s&solr_selection=&searchbox=&ms=0&page=%s&A=' % (catId.group(1), counter)
                    counter += 1
                    req = requests.get(moreProdsURL)
                    if req.status_code == 200:
                        counter += 1
                        pageSource = req.text
                        if 'no_selection' in pageSource:
                            break
                        if pageSource:
                            links = re.findall('href=."(.+?)"', pageSource)
                            if links:
                                for link in links:
                                    link = link.replace('\/', '/')
                                    link = link.replace('\\', '')
                                    if not 'http://www.poolsupplyworld.com' in link:
                                        link = 'http://www.poolsupplyworld.com' + link
                                    if not link in prodURLs:
                                        prodURLs.append(link)
                                        print 'prod URL::' + link
                                        looping = False
                    else:
                        looping = False
        #break#######################################################>>>>>>>>>>>>>> Remove this one.

#4.You have the full prods now, Just Scrape them using BS4

#Preparing CSV File
oFile = open('data.csv', "ab")
writer = csv.writer(oFile, delimiter=',', quotechar='"')
headersRow = ['SKU', 'ProductName', 'Category', 'SubCategories', 'Price', 'Description', 'FreeShipping', 'Ref']
writer.writerow(headersRow)

prodURLs = set(prodURLs)

for prodURL in prodURLs:
    sKU = ''
    productName = ''
    category = ''
    subCategories = ''
    price = ''
    description = ''
    freeShipping = ''

    req = requests.get(prodURL)
    if req.status_code == 200:
        pageSource = req.text
        if pageSource:
            soup = BeautifulSoup(pageSource)

            #SKU
            sKUTag = soup.find('meta', {'itemprop': 'sku'})
            if sKUTag:
                sKU = sKUTag.attrs['content']

            #productName
            productNameTag = soup.find('h1', {'itemprop': 'name'})
            if productNameTag:
                productName = productNameTag.text.replace('\n', '').strip()

            #category & subs
            cats = soup.find_all('div', {'class': 'crumb_text'})
            if cats:
                catCounter = 0
                for cat in cats:
                    catName = cat.text
                    catName = catName.replace('\n', '').strip()
                    subCategories = catName + ';' + subCategories
                    catArr = subCategories.split(';')
                if catArr:
                    category = catArr[0]
                    subCategories = subCategories.replace(category+';', '')
                    if subCategories.endswith(';'):
                        subCategories = subCategories[:-1]
            #Price
            priceDiv = soup.find('div', {'class': 'price_section'})
            if priceDiv:
                price = priceDiv.text.strip()
                price = price.replace('\n', '')

            #description
            descriptionTag = soup.find('div', {'id': 'description'})
            if descriptionTag:
                description = descriptionTag.text.replace('\n', '')
                description = description.replace('\r', '')
                description = description.replace('\t', '').strip()

            if 'id="badge_Free_Shipping"' in pageSource:
                freeShipping = 'Yes'
            else:
                freeShipping = 'No'

            if not sKU == '':
                newRow = [sKU.encode('utf-8'),
                          productName.encode('utf-8'),
                          category.encode('utf-8'),
                          subCategories.encode('utf-8'),
                          price.encode('utf-8'),
                          description.encode('utf-8'),
                          freeShipping.encode('utf-8'),
                          prodURL.encode('utf-8')]
                writer.writerow(newRow)
                print newRow

oFile.close()

print 'Done'