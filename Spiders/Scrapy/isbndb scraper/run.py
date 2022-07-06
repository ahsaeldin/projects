#!/usr/bin/python

from scrapy import cmdline

#IsASINABook
#ASINsExtractor -- from http://www.amazon.com/s/ref=sr_nr_p_n_feature_browse-b_mrr_0?fst=as%3Aoff&rh=n%3A283155%2Cp_n_feature_nine_browse-bin%3A3291437011%2Cp_n_feature_browse-bin%3A2656022011&bbn=283155&ie=UTF8&qid=1440785439&rnid=618072011
#ASINsExtractorFromBestSellerPages -- http://www.amazon.com/best-sellers-books-Amazon/zgbs/books

name = 'isbns'
if __name__ == '__main__':
    command = "scrapy crawl {0}".format(name).split()
    cmdline.execute(command)
