#!/usr/bin/python
# -*- coding: gbk -*-

import gensim
import jieba
import sys
import pandas as pd
import glob
import datetime
import csv
from itertools import izip_longest


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def timestamp_to_date(days):
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(days))


def item_similarity(item, item_pool):
    jieba.load_userdict(prefix_path + '/PG dict.txt')

    num_item = len(item_pool)

    class MyCorpus(object):
        def __iter__(self):
            for seg in item_pool:
                yield list(jieba.cut(seg, cut_all=False))

    corp = MyCorpus()
    dictionary = gensim.corpora.Dictionary(corp)
    corpus = [dictionary.doc2bow(text) for text in corp]
    lsi = gensim.models.LsiModel(corpus, id2word=dictionary, num_topics=num_item)
    query = list(jieba.cut(item, cut_all=False))
    vec_bow = dictionary.doc2bow(query)
    vec_lsi = lsi[vec_bow]
    index = gensim.similarities.MatrixSimilarity(lsi[corpus])
    sims = index[vec_lsi]
    sims = sorted(enumerate(sims), key=lambda x: -x[1])
    # under 0.9 should have manual check
    return list(sims[x][0] for x in range(0, num_item) if sims[x][1] > 0.7)


def get_dm_item_pool(dm_num):
    item_pool = list()
    # dm 1522/1561 have additional scorecard
    dm1522_path = prefix_path + '/DM Scorecard/WM a1522 1 DM Scorecard.csv'
    dm1561_path = prefix_path + '/DM Scorecard/WM a1561 1 DM Scorecard.csv'
    dict_path = prefix_path + '/DM Dictionary/WM %s UPC Dictionary.csv' % dm_num

    dict_df = pd.read_csv(dict_path, encoding='gbk')
    for cell_value in dict_df['Item ']:
        if cell_value not in item_pool:
            item_pool.append(str(cell_value))

    # append long term dm
    return item_pool


def gen_long_term_item_pool(y1516_dict_df, year, month1, month2):
    item_pool = list()

    max_date = datetime.date(year, month2, 1) + datetime.timedelta(days=31)
    min_date = datetime.date(year, month1, 1)

    df = y1516_dict_df[(y1516_dict_df['Month'] >= min_date) & (y1516_dict_df['Month'] < max_date)]
    for index, row in df.iterrows():
        if row[6] not in item_pool:
            item = row[6] + ' ' + row[7] + ' ' # + str(row[3]) # col: item desc 1/size/price
            item_pool.append(item)

    return item_pool


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    ignore_names = list([u'Unnamed', u'SFA Code', u'Store Name', u'Project Task', u'货架', u'端架', u'iAudit', u'卖进执行活动', u'形象堆头', u'陈列架', u'全国', u'门店'])
    ignore_cols = list([u'Unnamed', u'SFA Code', u'Store Name', u'Project Task'])
    dm_path = prefix_path + '/DM Scorecard/WM 1*.csv'
    y1516_dm_dict_path = prefix_path + '/DM Dictionary/WM Y1516 Long Term DM Dictionary.csv'

    y1516_dict_df = pd.read_csv(y1516_dm_dict_path, encoding='gbk')
    y1516_dict_df['Month'] = y1516_dict_df['Month'].apply(timestamp_to_date)

    jieba.load_userdict(prefix_path + '/PG dict.txt')

    for f in glob.glob(dm_path):
        # get year and month from
        dm_num = f.split()[-6]
        month1 = int(f.split()[-2])
        month2 = int(f.split()[-1].split('.')[0])
        year = int('20' + dm_num[0:2])

        item_pool = get_dm_item_pool(dm_num)
        item_pool += gen_long_term_item_pool(y1516_dict_df, year, month1, month2)

        total_cnt = 0
        matched_cnt = 0
        mapping_res = list()
        matched_q = list()
        unmatched_q = list()

        # mapping with dm scorecard
        dm_dict_df = pd.read_csv(f, encoding='gbk')
        for item in dm_dict_df.columns.values:
            if not any(s in item for s in ignore_names):
                total_cnt += 1
                res = item_similarity(item, item_pool)
                if any(res):
                    matched_cnt += 1
                    mapping_res.append([item, res])
                    matched_q.append(item)
                # for res in mapping_res:
                #     print item_pool[res].decode('gbk')
                # if not any(mapping_res):
                #     print 'no res'
                else:
                    unmatched_q.append(item)
            else:
                unmatched_q.append(item)

        # save res
        matched_items = list()
        with open(prefix_path + '/DM POS Mapping/WM %s UPC Dict Total %s Matched %s.csv' % (dm_num, total_cnt, matched_cnt), 'wb') as res_file:
            wr = csv.writer(res_file, quoting=csv.QUOTE_ALL)
            for item in mapping_res:
                for res in item[1]:
                    matched_items.append(item_pool[res])
                    wr.writerow([item[0], item_pool[res]])

        with open(prefix_path + '/DM POS Mapping/WM %s UPC Dict Total Potential Matched.csv' % dm_num, 'wb') as po_file:
            wr = csv.writer(po_file, quoting=csv.QUOTE_ALL)
            po_matched_q = list()
            for each_q in unmatched_q:
                if not any(s in each_q for s in ignore_cols):
                    po_matched_q.append(each_q)

            for item1, item2 in izip_longest(po_matched_q, (s for s in item_pool if s not in matched_items)):
                wr.writerow([item1, item2])

        # break # one loop test
        print dm_num + ' Done'

    print 'Done'

if __name__ == '__main__':
    main()