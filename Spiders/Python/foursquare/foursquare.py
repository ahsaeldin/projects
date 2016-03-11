import ujson
import requests
import foursquare

CLIENT_ID = u'YOUR CLIENT ID HERE'
CLIENT_SECRET = u'YOUR CLIENT SECRET HERE'

client = foursquare.Foursquare(client_id=CLIENT_ID,
                               client_secret=CLIENT_SECRET)

venues = client.venues.search(params={'near': 'chicago'})

url = client.oauth.auth_url()

for venue in venues[u'venues']:
    print venue

    venue_id = venue[u'id']
    photos = client.venues.photos(venue_id, params={'limit': '200'})
    if 'photos' in photos:
        photos = photos[u'photos']
        photos_count = photos[u'count']
        photos = photos[u'items']
        for photo in photos:
            print photo

    similar = client.venues.similar(venue_id)
    if 'similarVenues' in similar and similar[u'similarVenues'][u'count'] > 0:
        similar_count = similar[u'similarVenues'][u'count']
        similar_items = similar[u'similarVenues'][u'items']
        for similar_item in similar_items:
            print similar_item

    tips = client.venues.tips(venue_id)
    if 'tips' in tips and tips[u'tips'][u'count'] > 0:
        tips_count = tips[u'tips'][u'count']
        tips = tips[u'tips'][u'items']
        for tip in tips:
            print tip

    links = client.venues.links(venue_id)
    if 'links' in links and links[u'links'][u'count'] > 0:
        links_count = links[u'links'][u'count']
        links = links[u'links'][u'items']
        for link in links:
            print link

    menus = client.venues.menu(venue_id)
    if 'menu' in menus and menus[u'menu'][u'menus'][u'count'] > 0:
        menus_count = menus[u'menu'][u'menus'][u'count']
        menus = menus[u'menu'][u'menus'][u'items']
        for menu in menus:
            print menu

    lists = client.venues.listed(venue_id)
    if 'lists' in lists and lists[u'lists'][u'count'] > 0:
        lists_count = lists[u'lists'][u'count']
        groups = lists[u'lists'][u'groups']
        for group in groups:
            print group

    # extracting likes, hours and nextvenues using direct api requests because they are not supported in foursquare lib.
    not_supported = [u'likes', u'hours', u'nextvenues']
    for n in not_supported:
        x = 'https://api.foursquare.com/v2/venues/%s/%s?client_id=%s&client_secret=%s&v=20160311' % (venue_id,
                                                                                                     n,
                                                                                                     CLIENT_ID,
                                                                                                     CLIENT_SECRET)
        req = requests.get(x)
        if req.status_code == 200:
            data = ujson.loads(req.text)

            if u'likes' in data[u'response'] and data[u'response'][unicode(n)][u'count'] > 0:
                likes = data[u'response'][u'likes'][u'items']
                print likes
            elif u'hours' in data[u'response'] and len(data[u'response'][u'hours']) > 0:
                hours = data[u'response'][u'hours'][u'timeframes']
                popular = data[u'response'][u'popular'][u'timeframes']
                print hours
                print popular
            elif u'nextVenues' in data[u'response'] and data[u'response'][u'nextVenues'][u'count'] > 0:
                next_venues = data[u'response'][u'nextVenues'][u'items']
                for next_venue in next_venues:
                    print next_venue

