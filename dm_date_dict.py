#!/usr/bin/python
# -*- coding: gbk -*-

import xlrd
import sys
import pandas as pd

prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
cols_name = ['dm_num',
             'start_date',
             'end_date']


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    dm_sheet = xlrd.open_workbook(
        prefix_path + '/Raw Data/WM# Tab Performance -2016 汇总-1668.xlsx').sheet_by_name(u'1.New tracking work sheet')

    dm_sheet2 = xlrd.open_workbook(
        prefix_path + '/Raw Data/WM# Tab Performance -2015.xlsx').sheet_by_name(u'1.New tracking work sheet')

    dm_lst = list()
    res_df = pd.DataFrame(columns=cols_name)

    for row in range(1, dm_sheet.nrows):

        if isinstance(dm_sheet.row(row)[0].value, unicode):
            continue

        dm_num = int(dm_sheet.row(row)[0].value)

        if dm_num in dm_lst:
            continue

        else:
            dm_lst.append(dm_num)
            str_period = dm_sheet.row(row)[1].value
            print str_period
            start_date = str_period.split('-')[0]
            end_date = '2016.' + str_period.split('-')[1]

        res_df = res_df.append({'dm_num': dm_num,
                            'start_date': start_date,
                            'end_date': end_date},
                           ignore_index=True)

    for row in range(1, dm_sheet2.nrows):

        if isinstance(dm_sheet2.row(row)[0].value, unicode):
            continue

        dm_num = int(dm_sheet2.row(row)[0].value)

        if dm_num in dm_lst:
            continue

        else:
            dm_lst.append(dm_num)
            str_period = dm_sheet2.row(row)[1].value
            print str_period
            start_date = str_period.split('-')[0]
            end_date = '2015.' + str_period.split('-')[1]

        res_df = res_df.append({'dm_num': dm_num,
                                'start_date': start_date,
                                'end_date': end_date},
                               ignore_index=True)

        # break
    res_df.to_csv(prefix_path + '/Output/DM Date Dict.csv' )

    print 'Done'


if __name__ == '__main__':
    main()