import re
import csv
import random
import MyLines1
import requests
from bs4 import BeautifulSoup
from itertools import izip_longest

URLsFile = 'URLs.dat'
proxiesFile = 'proxies.dat'
dataFileName = 'RugsData.csv'
headers = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
           "Accept-Encoding": "gzip,deflate,sdch",
           "Accept-Language": "en-US,en;q=0.8",
           "User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.107 Safari/537.36"}

MyLines1.removeDataFile(dataFileName)
proxies = MyLines1.loadProxies(proxiesFile)

oFile = open(dataFileName, "ab")
writer = csv.writer(oFile, delimiter=',', quotechar='"', quoting=csv.QUOTE_ALL)

headersRow = ["sku",
              "product_type",
              "post_title",
              "price",
              "stock",
              "post_parent",
              "post_content",
              "variation",
              "var_color",
              "var_size",
              "var_rug-pad",
              "featured_image",
              "pa_cf-material",
              "pa_cf-weave",
              "pa_cf-origin",
              "product_gallery",
              "ref_url"]

writer.writerow(headersRow)

def addRow(soup, product_type, sku, post_parent, var_size, var_rug_pad, price, color):
    post_title = ''
    description = ''
    variation = 'Color->1->1|Size->1->1|Rug Pad->1->1'
    featured_image = ''
    pa_cf_material = ''
    pa_cf_weave = ''
    pa_cf_origin = ''
    product_gallery = ''

    #RefURL
    ref_url = prod

    if not post_parent == '':
        stock = '10'
    else:
        #Stock
        stock = ''

    #ProductName
    prodNameTag = soup.find('h2', id='pinName')
    if prodNameTag:
        post_title = prodNameTag.text.strip()

    #Description
    descTag = soup.find('p', id='productDescription')
    if descTag:
        description = descTag.text.replace(',', ';')
        description = description.replace('\r', '')
        description = description.replace('\n', '')
        description = description.replace('Description:', '')
        description = description.strip()

    #Image
    imageTag = soup.find('img', id='smallImageLoc')
    if imageTag:
        featured_image = imageTag.attrs['src']

    #Material
    featuresTag = soup.find_all('div', {'class': 'productFeatures'})
    for feature in featuresTag:
        text = feature.text
        if 'Material' in text:
            pa_cf_material = re.sub('\n|:|Material', '', text).strip()
        elif 'Weave' in text:
            pa_cf_weave = re.sub('\n|:|Weave', '', text).strip()

    #Origin
    bullLstTag = soup.find('ul', id='bulletList')
    if bullLstTag:
        bullLstTag = bullLstTag.find_all('li')
        for bullet in bullLstTag:
            text = bullet.text
            if 'Origin' in text:
                pa_cf_origin = re.sub('\n|:|Origin', '', text).strip()
                break

    #product_gallery
    product_galleryTag = soup.find('div', {'class': 'ad-thumbs'})
    if product_galleryTag:
        imagesLinksTags = product_galleryTag.find_all('a')
        for img in imagesLinksTags:
            imgHref = img.attrs['href']
            if not imgHref in product_gallery:
                if product_gallery == '':
                    product_gallery = imgHref
                else:
                    product_gallery = product_gallery + '|' + imgHref

    for feature in featuresTag:
        text = feature.text
        if 'Material' in text:
            pa_cf_material = re.sub('\n|:|Material', '', text).strip()
        elif 'Weave' in text:
            pa_cf_weave = re.sub('\n|:|Weave', '', text).strip()

    r0w = [sku.encode('utf-8'),
           product_type.encode('utf-8'),
           post_title.encode('utf-8'),
           price.encode('utf-8'),
           stock.encode('utf-8'),
           post_parent.encode('utf-8'),
           description.encode('utf-8'),
           variation.encode('utf-8'),
           color.encode('utf-8'),
           var_size.encode('utf-8'),
           var_rug_pad.encode('utf-8'),
           featured_image.encode('utf-8'),
           pa_cf_material.encode('utf-8'),
           pa_cf_weave.encode('utf-8'),
           pa_cf_origin.encode('utf-8'),
           product_gallery.encode('utf-8'),
           ref_url.encode('utf-8')]

    writer.writerow(r0w)

    print r0w.__repr__() + '\n' + '----------------------------------------------------------------------------------------'

prodURls = MyLines1.loadFileIntoLIst('URLs.dat')
for prod in prodURls:
    print "Scraping::" + prod
    if len(proxies) > 0:
        rndProxy = 'http://' + MyLines1.PickRandomProxy(proxies)
        req = requests.get(prod, headers=headers, proxies={'http': rndProxy})
    elif len(proxies) == 0:
        req = requests.get(prod, headers=headers)

    if req.status_code == 200:
        pageSource = req.text
        soup = BeautifulSoup(pageSource)
        #SKU
        skuTAG = soup.find('h3', {"itemprop": 'productId'})
        if skuTAG:
            sku = skuTAG.text.replace('Add to Wishlist', '')
            sku = sku.replace('Item #:', '')
            sku = sku.strip()

        addRow(soup, 'variation_master', sku, '', '', '', '', '')

        #Size
        sizeTag = soup.find('select', id='sizeList')
        if sizeTag:
            Size = sizeTag.text
            Size = re.sub('SELECT\sA\sSIZE', '', Size)
            Size = re.sub("\(Recommended\)", '', Size)
            Size = re.sub('\n', ';', Size)
            Size = re.sub(';;', '', Size)
            Size = Size.replace(r"\'", "'").strip()
            if Size.endswith(';'):
                Size = Size[:-1].strip()
            sizeArr = Size.split(';')

            varsCount = 0
            sizeCount = 0
            colorLst = []
            sizeLst = []
            prices = []
            for size in sizeArr:
                varsCount += 1
                sizeType = ''
                sizePrice = ''
                sizePriceMatcher = re.search('\((.+?)\)', size)
                if sizePriceMatcher:
                    sizePrice = sizePriceMatcher.group(1).strip()
                    sizePrice = sizePrice.replace('$', '')
                    sizePrice = sizePrice.replace(',', '')
                    if 'pack' in sizePrice:
                        print 'pack in price'
                    else:
                        prices.append(sizePrice)

                colorMatcher = re.match('[a-zA-Z\s]+', size)
                if colorMatcher:
                    colorName = colorMatcher.group(0).strip()
                    colorLst.append(colorName)
                    sizeType = size.replace(colorName, "").strip()
                sizeLst.append(sizeType)

        #RugPad
        padList = []
        rugPadTag = soup.find('select', id='padList')
        if rugPadTag:
            RugPad = rugPadTag.text
            RugPad = re.sub('\n', ';', RugPad)
            RugPad = re.sub('SELECT\sA\sPAD', '', RugPad)
            RugPad = re.sub("\(Recommended\);", '', RugPad)
            RugPad = re.sub(';;', '', RugPad)
            RugPad = RugPad.replace(r"\'", "'").strip()
            if RugPad.endswith(';'):
                RugPad = RugPad[:-1].strip()
            if RugPad.startswith(';'):
                RugPad = RugPad[1:].strip()

            rugsArr = RugPad.split(';')
            rugsCount = 0

        sizeLst = set(sizeLst)
        colorLst = set(colorLst)
        padList = set(rugsArr)
        prices = set(prices)

        varsCount = 0
        for sizeT in sizeLst:
            varsCount += 1
            addRow(soup,
                   'product_variation',
                   sku + '-v' + str(varsCount),
                   sku,
                   sizeT,
                   '',
                   random.choice(list(prices)),
                   '')

        for colorT in colorLst:
            varsCount += 1
            addRow(soup,
                   'product_variation',
                   sku + '-v' + str(varsCount),
                   sku,
                   '',
                   '',
                   random.choice(list(prices)),
                   colorT)

        for padT in padList:
            varsCount += 1
            addRow(soup,
                   'product_variation',
                   sku + '-v' + str(varsCount),
                   sku,
                   '',
                   padT,
                   random.choice(list(prices)),
                   '')

        varsCount = 0

oFile.close()
print 'Done'