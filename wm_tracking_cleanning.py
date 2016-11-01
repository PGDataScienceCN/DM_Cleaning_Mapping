#!/usr/bin/python
# -*- coding: gbk -*-

import xlrd
import glob
import csv
import sys
import pandas as pd
import Levenshtein

prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def check_summary_sheet (dir_path):
    for f in glob.glob(dir_path):
        workbook = xlrd.open_workbook(f)
        # check Summary sheet whether exist
        if 'Summary' not in workbook.sheet_names():
            print (f)
            print ('%s' % workbook.sheet_names())


def get_summary_sheet(dir_path):
    for path in glob.glob(dir_path):
        workbook = xlrd.open_workbook(path)
        summary_sheet = workbook.sheet_by_name('Summary')
        num_rows = summary_sheet.nrows

        dm_idx = path.split('/')[-1].find('#')
        dm_num = path.split('/')[-1][dm_idx+1:dm_idx+5]

        # ncols is not consistent
        print summary_sheet.ncols

        # for curr_row in range(2, num_rows - 1):
        #     print (summary_sheet.row_values(curr_row))

        with open(prefix_path + '/DM Dictionary/WM %s UPC Dictionary.csv' % dm_num, 'wb') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            for curr_row in range(2, num_rows-1):
                writer.writerow(summary_sheet.row_values(curr_row))

            print ('DM %s is done' % dm_num)


def compare_dm_scorecard(sc_path, dm_date):
    for path in glob.glob(sc_path):
        df = pd.read_csv(path, encoding='gbk')

        # load dm_sku_dictionary
        dm_num = path.split()[-2][0:4]
        print 'Start DM:' + dm_num
        dict_path = prefix_path + '/DM Dictionary/WM %s UPC Dictionary.csv' % dm_num
        dict_df = pd.read_csv(dict_path, encoding='gbk')

        matched = list()
        unmatched = list()
        matched_cnt = 0
        unmatched_cnt = 0
        ignore_cols = list([u'Unnamed', u'SFA Code', u'Store Name', u'Project Task'])
        dm_cnt = 0
        for col in df.columns.values:
            if any(s in col for s in ignore_cols):
                dm_cnt += 1

        # TODO: needs a brand dictionary and a category dictionary and a size dictionary
        # TODO: add factor value for DMs
        for col in df.columns.values:
            for item in dict_df[u'Item ']:
                l_ratio = Levenshtein.ratio(col.decode('gbk'), item)
                if l_ratio > 0.5:
                    # Done: add UPC code
                    matched_cnt += 1
                    for upc in dict_df[dict_df[u'Item '] == item][u'Upc']:
                        matched.append([col, item, upc, dm_date[0], dm_date[1]])

                else:
                    # list potential matched
                    unmatched.append([col, item, l_ratio, dm_date[0], dm_date[1]])

        # Save match and unmatch results
        unmatched.sort(key=lambda x: x[2], reverse=True)

        unmatched_result = list()
        # remove duplicates
        for lst in unmatched:
            if lst not in unmatched_result:
                # get potential cnt
                if lst[2] >= 0.2:
                    unmatched_cnt += 1
                unmatched_result.append(lst)

        # save potential matched list
        with open((prefix_path + '/DM Pos Mapping/WM %s UPC Potential Mapping Total %s Potential %s.csv' %
                (dm_num, dm_cnt, unmatched_cnt)), 'wb') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerows(unmatched_result)

        with open(prefix_path + '/DM Pos Mapping/WM %s UPC Mapping Total %s Matched %s.csv' %
                (dm_num, dm_cnt, matched_cnt), 'wb') as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerows(matched)

        # for one loop test
        # break


def get_dm_date(sc_path):
    dm_date = sc_path.split()[-1].split('-')
    yr_num = sc_path.split()[-2][0:2]
    start_date = '20' + yr_num + '.' + dm_date[0]
    end_date = '20' + yr_num + '.' + dm_date[1].replace('.csv', '')
    # print start_date
    # print end_date
    return list([start_date.replace('.', '/'), end_date.replace('.', '/')])


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    dir_path = prefix_path + '/Raw Data/WM Tracking/WM DM #1501*.xlsx'
    t_path = prefix_path + '/DM Scorecard/WM 1501 1.1-1.14.csv'
    sc_path = prefix_path + '/DM Scorecard/WM 1*.csv'

    dm_date = get_dm_date(t_path)

    # check_summary_sheet(dir_path) #one time job
    # get_summary_sheet(dir_path)

    # 1 time checking, if AE put wrong or nick name, cannot be matched
    compare_dm_scorecard(sc_path, dm_date)
    print 'Done'

if __name__ == '__main__':
    main()