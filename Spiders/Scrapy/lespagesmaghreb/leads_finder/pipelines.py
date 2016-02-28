# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: http://doc.scrapy.org/en/latest/topics/item-pipeline.html

import os
import unicodecsv as csv
from sqlalchemy.orm import sessionmaker
from sqlalchemy import Column, Integer, String
from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()

class Lead(Base):

    __tablename__ = 'leads'

    Book_ID = Column(String, primary_key=True)
    Amazon_Best_Sellers_Rank = Column(Integer)
    ISBN_10 = Column(String)
    ISBN_13 = Column(String)
    ASIN = Column(String)
    Amazon_title = Column(String)
    Edition = Column(String)
    cover_type = Column(String)
    Amazon_author = Column(String)
    Amazon_publisher = Column(String)
    publication_date = Column(String)
    publication_year = Column(String)
    shipping_weight = Column(String)
    Amazon_language = Column(String)
    TimeStamp = Column(String)
    pages_length = Column(String)
    is_excluded = Column(Integer)
    exclude_reason = Column(String)

    def __repr__(self):
        return "<Lead(ListingID='%s', Name='%s)>" % (self.ListingID, self.Name)

class LeadFinSqlAlchemyPipeline(object):
    def __init__(self):
        self.meta = None
        self.engine = None

    def open_spider(self, spider):
        self.engine = engine = create_engine('sqlite:///leads.sqlite', echo=False)
        self.DBSession = sessionmaker(bind=engine)
        self.session = self.DBSession()

    def process_item(self, item, spider):
        is_updated = False

        Book_ID = item['Book_ID']
        is_added_before = self.session.query(Lead).filter_by(Book_ID=Book_ID).first()
        if not is_added_before:
            new_lead = Lead(Book_ID=Book_ID,
                            TimeStamp=item['TimeStamp'],
                            Amazon_Best_Sellers_Rank=item['Amazon_Best_Sellers_Rank'],
                            ISBN_10=item['ISBN_10'],
                            ISBN_13=item['ISBN_13'],
                            ASIN=item['ASIN'],
                            Amazon_title=item['Amazon_title'],
                            Edition=item['Edition'],
                            cover_type=item['cover_type'],
                            Amazon_author=item['Amazon_author'],
                            pages_length=item['pages_length'],
                            Amazon_publisher=item['Amazon_publisher'],
                            publication_date=item['publication_date'],
                            publication_year=item['publication_year'],
                            shipping_weight=item['shipping_weight'],
                            Amazon_language=item['Amazon_language'],
                            is_excluded=item['is_excluded'],
                            exclude_reason=item['exclude_reason'])
            try:
                self.session.add(new_lead)
                self.session.commit()
                is_updated = True
            except:
                print 'asdf'
        else:

            is_added_before.TimeStamp = item['TimeStamp']
            is_added_before.Amazon_Best_Sellers_Rank = item['Amazon_Best_Sellers_Rank']
            is_added_before.ISBN_10 = item['ISBN_10']
            is_added_before.ISBN_13 = item['ISBN_13']
            is_added_before.ASIN = item['ASIN']
            is_added_before.Amazon_title = item['Amazon_title']
            is_added_before.Edition = item['Edition']
            is_added_before.cover_type = item['cover_type']
            is_added_before.Amazon_author = item['Amazon_author']
            is_added_before.pages_length = item['pages_length']
            is_added_before.Amazon_publisher = item['Amazon_publisher']
            is_added_before.publication_date = item['publication_date']
            is_added_before.shipping_weight = item['shipping_weight']
            is_added_before.Amazon_language = item['Amazon_language']
            is_added_before.is_excluded = item['is_excluded']
            is_added_before.exclude_reason = item['exclude_reason']
            try:
                self.session.commit()
            except:
                print ''
            is_updated = True

        if not is_updated:
            print ''

        return item

    # def close_spider(self, spider):
    #     self.DBSession.clo
    #     self.ds.close()

#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################

#old pipeline
class YPPipeline(object):

    outputs_file_name = 'YP__Organized.csv'

    def process_item(self, item, spider):
        file_exists = False
        if os.path.exists(self.outputs_file_name):
            file_exists = True

        headers = ['TimeStamp',
                   'ZipCode',
                   'Keyword',
                   'Name',
                   'Phone',
                   'Categories',
                   'Website',
                   'Address',
                   'ListingURL',
                   'OverviewPageURL']

        test_file = open(self.outputs_file_name, 'ab')
        csvwriter = csv.DictWriter(test_file, quotechar='"', delimiter=',', fieldnames=headers)
        if not file_exists:
            csvwriter.writeheader()
        csvwriter.writerow(item)
        test_file.close()

        return item
