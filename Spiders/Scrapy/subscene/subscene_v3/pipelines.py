# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html

import os
import unicodecsv as csv
import subscene_v3.settings

class subscene_v3_pipleline(object):

    outputs_file_name = 'subs_Organized.csv'

    def process_item(self, item, spider):
        file_exists = False
        if os.path.exists(self.outputs_file_name):
            file_exists = True

        headers = ['TimeStamp',
                   'Movie_IMDB_ID',
                   'Movie_Name',
                   'Rating',
                   'RatingUsersCount',
                   'HI',
                   'Lang',
                   'Owner',
                   'OwnerURL',
                   'ListingURL',
                   'SubtitleURL',
                   'SourceURL']

        test_file = open(self.outputs_file_name, 'ab')
        csvwriter = csv.DictWriter(test_file, quotechar='"', delimiter=',', fieldnames=headers)
        if not file_exists:
            csvwriter.writeheader()
        csvwriter.writerow(item)
        test_file.close()

        subscene_v3.settings.remaining_urls.remove(item['SourceURL'])
        print str(len(subscene_v3.settings.remaining_urls)) + ' Remaining'

        return item
