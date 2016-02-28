# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html

import os
import moxa.settings
import unicodecsv as csv

class moxaPipeline(object):

    outputs_file_name = 'Moxa_Outs_Organized.csv'

    def process_item(self, item, spider):
        file_exists = False
        if os.path.exists(self.outputs_file_name):
            file_exists = True

        headers = ['SKU', 'MainCategory', 'Subcategory1', 'Subcategory2', 'Subcategory3', 'Subcategory4',
                   'Images', 'Datasheet', 'FeaturesAndBenefits', 'Overview', 'AvailableModels', 'CompatibleModules',
                   'HardwareSpecifications', 'Weight', 'Dimensions', 'RelatedProducts']

        test_file = open(self.outputs_file_name, 'ab')
        csvwriter = csv.DictWriter(test_file, quotechar='"', delimiter=',', fieldnames=headers)
        if not file_exists:
            csvwriter.writeheader()
        csvwriter.writerow(item)
        test_file.close()

        #moxa.settings.remaining_urls.remove(item['RefUrl'])
        #print str(len(moxa.settings.remaining_urls)) + ' Remaining'

        return item
