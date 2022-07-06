import re
import urllib
import hashlib
import mechanize
from bs4 import BeautifulSoup
import xml.etree.cElementTree as ET

def reportFailedURL(URL):
    print "Failed::" + URL
    f = open('failed-links.txt', 'a')
    f.write(URL + '\n')
    f.close()

URLs = []

with open('recipes.txt') as f:
    URLs = f.read().splitlines()

br = mechanize.Browser()

root = ET.Element('root')

count = 0
for URL in URLs:
    count += 1

    doc = ET.SubElement(root, "recipe")

    try:
        br.open(URL)
    except:
        reportFailedURL(URL)
        continue

    pageSource = br.response().read()

    soup = BeautifulSoup(pageSource)

    Site = URL.strip()

    Ingredient_tag = []
    ingredients = soup.find_all("span",attrs={"class":"ingredient"})
    if ingredients:
        ingredientsLst = [ingredient.text for ingredient in ingredients]
        for ingred in ingredientsLst:
            Ingredient_tag.append(ingred.strip())

    Procedure_tag = []
    procedures = soup.find_all("div",attrs={"class":"procedure-text"})
    if procedures:
        try:
            procedureLst = [procedure.text for procedure in procedures]
        except:
            reportFailedURL(URL)
            continue

        for proc in procedureLst:
            proc = proc.replace(';', ' ')
            proc = proc.replace(',', ' ')
            Procedure_tag.append(proc)

    titleNode = soup.find('title')
    if titleNode:
        title_tag = titleNode.text.strip()

    recipe_image = ''
    imgNode = soup.find('meta', attrs={"property":"og:image"})
    if imgNode:
        recipe_image = imgNode.attrs['content'].strip()

    recipe_image1 = ''
    imgNode1 = soup.find('img', attrs={"width":"518"})
    if imgNode1:
        recipe_image1 = imgNode1.attrs['src'].strip()

    if recipe_image1:
        recipe_image = recipe_image1

    #imageHash = ''
    #if recipe_image:
    #    imageHash = hashlib.md5(recipe_image).hexdigest() + '.jpg'
    #    imageHash = unicode(imageHash.decode("utf-8"))
    #    urllib.urlretrieve(recipe_image, "images/" + imageHash)

    photographerNode = soup.find('meta', attrs={"name":"description"})
    if photographerNode:
        content = photographerNode.attrs['content']
        matchObj = re.search('\[Photograph:(.+?)\]', content)
        if matchObj:
            photographer = matchObj.group(1).strip()
        else:
            photographer = ''

    tags = ''
    tagsNode = soup.find('ul', attrs={"class": "tags"})
    if tagsNode:
        children = tagsNode.findChildren()
        for child in children:
            if not child.text in tags:
                if tags == '':
                    tags = child.text.strip()
                else:
                    tags = tags + ',' + child.text.strip()

    yield_ = ''
    yieldNode = soup.find('span', attrs={"class":"info yield"})
    if yieldNode:
        yield_ = yieldNode.text.strip()

    active_time = ''
    activeTimeNode = soup.find('span', attrs={"class": "info prepTime"})
    if activeTimeNode:
        active_time = activeTimeNode.text.strip()

    total_time = ''
    totalTimeNode = soup.find('span', attrs={"class": "info totalTime duration"})
    if totalTimeNode:
        total_time = totalTimeNode.text.strip()

    special_equipment = ''
    matchObj = re.search('Special\sequipment:</span></td><td><span\sclass="info">(.+?)</span></td></tr>', pageSource)
    if matchObj:
        special_equipment = matchObj.group(1)

    specialEquipmentNode = soup.find_all('tr', text=re.compile('.+?/Special\sequipment/.+?'))
    if specialEquipmentNode:
        for spec in specialEquipmentNode:
            special_equipment = specialEquipmentNode.text.strip()

    desc = ''
    matchObj = re.search('<Attribute\sname="description">(.+?)</Attribute>', pageSource)
    if matchObj:
        desc = matchObj.group(1).strip()
    else:
        desc = ''

    author = ''
    matchObj = re.search('<Attribute\sname="author">(.+?)</Attribute>', pageSource)
    if matchObj:
        author = matchObj.group(1).strip()

    rating = ''
    matchObj = re.search('<Attribute\sname="recipe_rating">(.+?)</Attribute>', pageSource)
    if matchObj:
        rating = matchObj.group(1).strip()
    else:
        rating = 'NULL'

    recipeSource = ET.SubElement(doc, 'recipe_source')
    recipeSource.text = 'www.SeriousEats.com'

    recipeURL = ET.SubElement(doc, 'recipe_url')
    recipeURL.text = URL

    uniqueId = ET.SubElement(doc, 'unique_id')
    uniqueId.text = str(count)

    dish = ET.SubElement(doc, 'dish')
    titleArr = title_tag.split("|")
    if len(titleArr) > 0:
        dish.text = titleArr[0]
    else:
        dish.text = unicode(title_tag.decode("utf-8"))

    description = ET.SubElement(doc, 'description')
    description.text = unicode(desc.decode("utf-8"))

    categories = ET.SubElement(doc, 'categories')
    categories.text = 'NULL'

    authorNode = ET.SubElement(doc, 'recipe_author')
    authorNode.text = unicode(author.decode("utf-8"))

    recipeImage = ET.SubElement(doc, 'recipe_image')
    recipeImage.text = recipe_image

    photographerNode = ET.SubElement(doc, 'photo_credit')
    photographerNode.text = photographer

    activeTimeNode = ET.SubElement(doc, 'prep_time')
    activeTimeNode.text = active_time

    totalTimeNode = ET.SubElement(doc, 'total_time')
    totalTimeNode.text = total_time

    recipeCookTimeNode = ET.SubElement(doc, 'recipe_cook_time')
    recipeCookTimeNode.text = 'NULL'

    specialEquipmentNode = ET.SubElement(doc, 'special equipment')
    specialEquipmentNode.text = unicode(special_equipment.decode("utf-8"))

    ratingNode = ET.SubElement(doc, 'rating')
    ratingNode.text = rating

    servings = ET.SubElement(doc, 'servings')
    matchYield = re.sub('[A-Za-z]+', '', yield_).strip()
    if matchYield:
        servings.text = matchYield
    else:
        servings.text = yield_

    recipe_calories = ET.SubElement(doc, 'recipe_calories')
    recipe_calories.text = 'NULL'

    ingredients = ET.SubElement(doc, 'ingredients')
    for ingredient in Ingredient_tag:
        ingredientNode = ET.SubElement(ingredients, 'ingredient')
        ingredientNode.text = ingredient

    procs = ET.SubElement(doc, 'directions')
    for proc in Procedure_tag:
        procNode = ET.SubElement(procs, 'step')
        procNode.text = proc

    tagsNode = ET.SubElement(doc, 'tags')
    tagsNode.text = tags

    tree = ET.ElementTree(root)

    tree.write("data.xml", "utf-8")

    print str(count)

print 'Done'
