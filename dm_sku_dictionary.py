#!/usr/bin/python
# -*- coding: gbk -*-
import pandas as pd
import csv
import glob
import datetime
import xlrd
import sys
import jieba
import gensim
from itertools import izip_longest


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
dm_q_folder = prefix_path + '/Raw Data/DM1516/WM 1*.xlsx'
tab15_path = prefix_path + '/Raw Data/WM# Tab Performance -2015.xlsx'
tab16_path = prefix_path + '/Raw Data/WM# Tab Performance -2016.xlsx'
res_path = prefix_path + '/DM POS Mapping/'


def timestamp_to_date(days):
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(days))


# LSI model threshold 0.9
def item_similarity(item, item_pool, thres=0.9):
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
    return list(sims[x][0] for x in range(0, num_item) if sims[x][1] > thres)


# Warning: need reverse the format while mapping back
def pampers_reformat(org_str):
    nick_name1 = u'白帮'
    nick_name2 = u'绿帮'
    if nick_name1 in org_str:
        return org_str.replace(nick_name1, u'帮宝适特级棉柔（白帮）')
    elif nick_name2 in org_str:
        return org_str.replace(nick_name2, u'帮宝适超薄干爽(绿帮)')
    else:
        return org_str


def all_unique(lst):
    seen = list()
    return not any(i in seen or seen.append(i) for i in lst)


def get_dm_date(f):
    year = int('20' + f.split()[-2][0:2])
    month1 = int(f.split()[-1].split('.')[0])
    month2 = int(f.split()[-1].split('-')[1].split('.')[0])
    day1 = int(f.split()[-1].split('.')[1].split('-')[0])
    day2 = int(f.split()[-1].split('.')[-2])
    start_date = datetime.date(year, month1, day1)
    end_date = datetime.date(year, month2, day2)
    return start_date, end_date


def get_dm_lst(dm_num):

    file_path = prefix_path + '/Raw Data/DM1516/WM %s*.xlsx' % dm_num
    for path in glob.glob(file_path):
        workbook = xlrd.open_workbook(path)
    summary_sheet = workbook.sheet_by_name(workbook.sheet_names()[0])
    dm_lst = summary_sheet.row_values(2)
    dm_last_idx = dm_lst.index(u'DM Compliance Rate')
    del dm_lst[dm_last_idx:]
    dm_lst = filter(None, dm_lst)
    for i in range(len(dm_lst)):
        dm_lst[i] = pampers_reformat(dm_lst[i])

    # double check with details sheet
    details_sheet = workbook.sheet_by_name(u'问卷明细')
    if not any(dm in details_sheet.row_values(3) for dm in dm_lst):
        if all_unique(dm_lst):
            print 'DM: '+ dm_num + ' unmatched'

    return dm_lst


def get_dm_pool(dm_num, tab_dict):
    dm_pool = list()
    for row in tab_dict:
        if int(row[0]) == int(dm_num):
            dm_pool.append([row[1], row[2], row[3]])  # promo desc + upc +barcode
    return dm_pool


def gen_dm_dict(dm_num, tab15_dict, tab16_dict):
    dm_lst = get_dm_lst(dm_num)
    total_cnt = len(dm_lst)
    matched_cnt = 0
    res_lst = list()
    un_matched_lst = list()

    if dm_num < '1600':
        dm_pool = get_dm_pool(dm_num, tab15_dict)
    else:
        dm_pool = get_dm_pool(dm_num, tab16_dict)

    un_matched_pool = list(dm_pool)

    for dm in dm_lst:
        res_index = item_similarity(dm, [s[0] for s in dm_pool])
        if any(res_index):
            matched_cnt += 1
        else:
            un_matched_lst.append(dm)
        for i in res_index:
            try:
                un_matched_pool.remove(dm_pool[i])
            except ValueError:
                pass
            res_lst.append([dm.strip(), dm_pool[i][0].strip(), dm_pool[i][1], dm_pool[i][2]])

    # mapping and save to csv
    with open(res_path + 'DM %s UPC Dictionary Total %s Matched %s.csv' % (dm_num, total_cnt, matched_cnt), 'wb') as f:
        wr = csv.writer(f, quoting=csv.QUOTE_ALL)
        for row in res_lst:
            wr.writerow(row)

    # save manual mapping res
    if any(un_matched_pool):
        with open(res_path + 'DM %s Potential Matched %s.csv' % (dm_num, total_cnt-matched_cnt), 'wb') as f:
            wr = csv.writer(f, quoting=csv.QUOTE_ALL)
            for item1, item2 in izip_longest(un_matched_lst, un_matched_pool):
                wr.writerow([item1, item2[0], item2[1], item2[2]])
    return 0


def gen_edlp_dict(dm_num, start_date, end_date):
    return 0


def get_tab15_dict(tab_sheet):
    tab_dict = list()
    for row in range(tab_sheet.nrows):
        try:
            tab_num = tab_sheet.row_values(row)[0]
            int(tab_num)
        except ValueError:
            continue  # skip some rows
        tab_dict.append([tab_num,
                         tab_sheet.row_values(row)[4],  # promo desc
                         tab_sheet.row_values(row)[6],  # upc
                         tab_sheet.row_values(row)[7],  # barcode
                         tab_sheet.row_values(row)[8]])   # primary desc
    return tab_dict


def get_tab16_dict(tab_sheet):
    tab_dict = list()
    for row in range(tab_sheet.nrows):
        try:
            tab_num = tab_sheet.row_values(row)[0]
            int(tab_num)
        except ValueError:
            continue  # skip some rows
        tab_dict.append([tab_num,
                         tab_sheet.row_values(row)[4],  # promo desc
                         tab_sheet.row_values(row)[6],  # upc
                         tab_sheet.row_values(row)[7]])  # barcode
    return tab_dict


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    jieba.load_userdict(prefix_path + '/PG dict.txt')
    # load tab performance
    tab15_sheet = xlrd.open_workbook(tab15_path).sheet_by_name('1.New tracking work sheet')
    tab15_dict = get_tab15_dict(tab15_sheet)
    tab16_sheet = xlrd.open_workbook(tab16_path).sheet_by_name('1.New tracking work sheet')
    tab16_dict = get_tab16_dict(tab16_sheet)

    dm_date_dict = list()

    # Todo: DM 1561 needs exception cleaning
    for f in glob.glob(dm_q_folder):
        if f.split()[-3] == '2':
            # Todo: manually cleaning
            continue
        else:
            dm_num = f.split()[-2][0:4]
            start_date, end_date = get_dm_date(f)

            # one time job
            # dm_date_dict.append([dm_num, start_date, end_date])
            # with open(prefix_path + 'dm_date_mapping.csv', 'wb') as f:
            #     wr = csv.writer(f, quoting=csv.QUOTE_ALL)
            #     for i in dm_date_dict:
            #         wr.writerow(i)

            gen_dm_dict(dm_num, tab15_dict, tab16_dict)  # save to csv
            # created a new script
            # gen_edlp_dict(dm_num, start_date, end_date)  # save to csv

            print dm_num

        # break  # one loop test

    print 'Done'


if __name__ == '__main__':
    main()