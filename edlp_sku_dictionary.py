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
import NO_USE_swf_edlp_pdq_cleaning
import dm_sku_dictionary


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
dm_q_folder = prefix_path + '/Raw Data/DM1516/WM 1*.xlsx'
edlp_path = prefix_path + '/Raw Data/SWF_EDLP_PDQ SKU.xlsx'
res_path = prefix_path + '/EDLP POS Mapping/'
ignore_names = list([u'Unnamed',
                     u'SFA Code',
                     u'Store Name',
                     u'Project Task',
                     u'货架',
                     u'端架',
                     u'iAudit',
                     u'卖进并执行活动',
                     u'形象堆头',
                     u'陈列架',
                     u'档条',
                     u'全国'])


def float_to_date(float_date):
    return datetime.datetime.strptime(str(int(float_date)), '%Y%m%d').date()


def get_edlp_others_lst(f):
    dm_book = xlrd.open_workbook(f)
    summary_sheet = dm_book.sheet_by_name(dm_book.sheet_names()[0])
    details_sheet = dm_book.sheet_by_name(u'问卷明细')
    edlp_q_lst = summary_sheet.row_values(2)
    others_q_lst = list()

    try:
        edlp_last_idx = edlp_q_lst.index(u'EDLP Pack Compliance Rate')
        edlp_start_idx = summary_sheet.row_values(1).index(u'EDLP Pack Tracking')
        del edlp_q_lst[edlp_last_idx:]
        del edlp_q_lst[:edlp_start_idx]
        for i in range(len(edlp_q_lst)):
            edlp_q_lst[i] = dm_sku_dictionary.pampers_reformat(edlp_q_lst[i])
    except ValueError:
        edlp_q_lst = list()
        pass

    # add details sheet dm
    for dm_item in details_sheet.row_values(3):
        if dm_item not in summary_sheet.row_values(2):
            if not any(s in dm_item for s in ignore_names):
                others_q_lst.append(dm_item)

    return edlp_q_lst, others_q_lst


# computing consuming, will optimize later
def get_item_lst(raw_sheet, start_date, end_date):
    item_lst = list()
    for i in range(1, raw_sheet.nrows):
        row_value = raw_sheet.row_values(i)
        if start_date <= dm_sku_dictionary.timestamp_to_date(row_value[10]) <= end_date:
            # unique
            if not any(row_value[1] == s[0] for s in item_lst):
                item_lst.append([row_value[1],  # 0 upc
                                 row_value[2],  # 1 barcode
                                 row_value[3],  # 2 cate
                                 row_value[4],  # 3 price
                                 dm_sku_dictionary.pampers_reformat(row_value[7]),  # 4 item desc
                                 row_value[8],  # 5 size
                                 row_value[9],  # 6 swf/tab/pdq/swf
                                 row_value[5],  # 7 start date
                                 row_value[6]])  # 8 end date
    return item_lst


def gen_edlp_dict(dm_num, raw_sheet, start_date, end_date, f):
    matched_cnt = 0
    unmatched_q_lst = list()
    res_lst = list()

    edlp_q_lst, others_lst = get_edlp_others_lst(f)
    item_lst = get_item_lst(raw_sheet, start_date, end_date)

    unmatched_item_lst = list(item_lst)

    if edlp_q_lst:
        for edlp in edlp_q_lst:
            edlp_res_idx = dm_sku_dictionary.item_similarity(edlp, [s[4] for s in item_lst], thres=0.5)
            if any(edlp_res_idx):
               matched_cnt += 1
            else:
                unmatched_q_lst.append(edlp)
            for i in edlp_res_idx:
                try:
                    unmatched_item_lst.remove(item_lst[i])
                except ValueError:
                    pass
                res_lst.append([edlp, item_lst[i]])

    # save to csv
    with open(res_path + 'DM %s EDLP Dict Total %s Matched %s.csv' % (dm_num, len(edlp_q_lst), matched_cnt), 'wb') as f:
        wr = csv.writer(f, quoting=csv.QUOTE_ALL)
        for row in res_lst:
            res = [row[0]] + row[1]
            wr.writerow(res)

    if any(unmatched_item_lst):
        with open(res_path + 'DM %s Unmatched CNT %s.csv' % (dm_num, len(unmatched_item_lst)), 'wb') as f:
            wr = csv.writer(f, quoting=csv.QUOTE_ALL)
            for item1, item2 in izip_longest(unmatched_q_lst, unmatched_item_lst):
                res = [item1] + item2
                wr.writerow(res)

    if any(others_lst):
        with open(res_path + 'DM %s Others.csv' % dm_num, 'wb') as f:
            for item in others_lst:
                f.write(item + '\n')

    return 0


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    jieba.load_userdict(prefix_path + '/PG dict.txt')
    # load edlp performance
    workbook, raw_sheet = NO_USE_swf_edlp_pdq_cleaning.load_sheet(edlp_path)

    # mapping with dm
    # Todo: DM 1561 needs exception cleaning
    for f in glob.glob(dm_q_folder):
        if f.split()[-3] == '2':
            # Todo: manually cleaning
            continue
        else:
            dm_num = f.split()[-2][0:4]
            start_date, end_date = dm_sku_dictionary.get_dm_date(f)
            gen_edlp_dict(dm_num, raw_sheet, start_date, end_date, f)  # save to csv
            print dm_num

        # break  # one loop test

    print 'Done'


if __name__ == '__main__':
    main()