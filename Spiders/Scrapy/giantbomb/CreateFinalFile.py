import ast
import ujson
import unicodecsv as csv

def saveOutput(data, keys, filename):
    file = open(filename, 'ab')
    dict_writer = csv.DictWriter(file, keys, delimiter=',', quoting=csv.QUOTE_ALL)
    if len(data) == 0:
        dict_writer.writeheader()
    else:
        dict_writer.writerows(data)

urls = dict()
original_urls_file = 'games_urls.csv'
csv_file = csv.DictReader(open(original_urls_file, 'rb'), delimiter=',')
countter = 0
for line in csv_file:
    countter += 1
    print str(countter)
    game_url = line['game_url']
    overview = line['overview']
    urls[game_url] = overview

keys_set = set()
keys_set.add(u'TimeStamp')
keys_set.add(u'Game_ID')
keys_set.add(u'Name')
keys_set.add(u'CoverArt')
keys_set.add(u'Desc')
keys_set.add(u'OverViewRef')
keys_set.add(u'RefURL')
keys_set.add(u'FileName')

final_lst = list()

csv_file = csv.DictReader(open('Outputs.csv', 'rb'), delimiter=',')
for line in csv_file:

    print line

    final_dict = dict()
    final_dict[u'TimeStamp'] = line['TimeStamp']
    final_dict[u'Game_ID'] = line['Game_ID']
    final_dict[u'Name'] = line['Name']
    final_dict[u'CoverArt'] = line['CoverArt']
    final_dict[u'Desc'] = line['Desc']
    final_dict[u'OverViewRef'] = line['OverViewRef']
    final_dict[u'RefURL'] = line['RefURL']
    final_dict[u'NoOfReleases'] = line['NoOfReleases']

    urls.pop(line['RefURL'])  # the remaining non poped out is the failed games urls to re-check later

    try:
        if '/xbox-one/' in line['OverViewRef']:
            final_dict[u'FileName'] = 'xbox-one.csv'
        elif '/nintendo-entertainment-system/' in line['OverViewRef']:
            final_dict[u'FileName'] = 'nintendo-entertainment-system.csv'
        elif '/playstation-4/' in line['OverViewRef']:
            final_dict[u'FileName'] = 'playstation-4.csv'
    except:
        continue

    game_details = ast.literal_eval(line['GameDetails'])
    for k, v in game_details.iteritems():
        final_dict[k] = v
        keys_set.add(k)

    final_lst.append(final_dict)

koko_list = set()
keys_set = list(keys_set)
keys_set.sort()
keys_set = ['TimeStamp',
            'Game_ID',
            'Name',
            'NoOfReleases',
            'Aliases',
            'CoverArt',
            'Desc',
            'Developer',
            'First release date',
            'Franchises',
            'Genre',
            'Platform',
            'Publisher',
            'Rating',
            'Release Date',
            'Theme',
            'RefURL',
            'OverViewRef',
            'FileName']

for l in final_lst:
    print l
    resultset = list()
    resultset.append(l)
    if l['FileName'] not in koko_list:
        koko_list.add(l['FileName'])
        saveOutput([], keys_set, l['FileName'])
    saveOutput(resultset, keys_set, l['FileName'])


