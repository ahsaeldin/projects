import urllib
import hashlib
from urlparse import urlparse
from os.path import splitext, basename

with open('images.txt') as f:
    URLs = f.read().splitlines()

count = 0
all = str(len(URLs))
for image_url in URLs:
    count += 1
    print str(count) + ' of ' + all
    o = urlparse(image_url)
    filename, file_ext = splitext(basename(o.path))
    urllib.urlretrieve(image_url, "images/" + filename + file_ext)

    #imageHash = hashlib.md5(recipe_image).hexdigest() + '.jpg'
    #imageHash = unicode(imageHash.decode("utf-8"))
    #urllib.urlretrieve(recipe_image, "images/" + imageHash)