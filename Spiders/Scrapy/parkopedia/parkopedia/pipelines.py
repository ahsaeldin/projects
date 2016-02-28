# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html

import os
import unicodecsv as csv
import parkopedia.settings

class ParkoPediaPipeline(object):

    outputs_file_name = 'Parkopedia_Outs_Organized.csv'

    def process_item(self, item, spider):
        file_exists = False
        if os.path.exists(self.outputs_file_name):
            file_exists = True

        headers = ['ParkURL', 'TimeStamp', 'Country', 'State', 'City', 'Neighborhood', 'PostCode', 'Latitude', 'Longitude',
                   'ParkID', 'ParkingType', 'ParkName', 'NumberOfSpaces', 'Address', 'Phone', 'PaymentsAccepted',
                   'NumberofParkingSpaces', 'Type', 'Features', 'Info', 'Progress']

        city = item['City']
        test_file = open(self.outputs_file_name, 'ab')
        csvwriter = csv.DictWriter(test_file, quotechar='"', delimiter=',', fieldnames=headers)
        if not file_exists:
            csvwriter.writeheader()
        csvwriter.writerow(item)
        test_file.close()

        print str(item['Progress']) + ' / ' + str(len(parkopedia.settings.remaining_cities))

        return item
