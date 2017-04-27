#!/usr/bin/python
# -*- coding: gbk -*-

import xlrd
import sys
import pandas as pd
from datetime import date, datetime, timedelta

prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
cols_name = ['dm_num',
             'dm_date',
             'day_of_dm',
             'dm_days',
             'start_date',
             'end_date']


def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + timedelta(n)


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    dm_csv = pd.read_csv('C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Knime Opt/dm date dict2.csv')

    print dm_csv.columns
    res_df = pd.DataFrame(columns=cols_name)

    # dm_csv['start_date'] = pd.to_datetime(dm_csv['start_date'], format='%Y/%m/%d')
    # dm_csv['end_date'] = pd.to_datetime(dm_csv['end_date'], format='%Y/%m/%d')

    for dm in dm_csv['dm_num'].unique():
        start_date = dm_csv[dm_csv['dm_num'] == dm]['start_date'].item()
        end_date = dm_csv[dm_csv['dm_num'] == dm]['end_date'].item()

        start_date = datetime.strptime(start_date, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date, '%Y-%m-%d').date()

        day_of_dm = 1
        for dm_date in daterange(start_date, end_date):
            res_df = res_df.append({'dm_num': dm,
                                     'dm_date': dm_date,
                                     'day_of_dm': day_of_dm,
                                     'dm_days': dm_csv[dm_csv['dm_num'] == dm]['DM Days'].item(),
                                     'start_date': start_date,
                                     'end_date': end_date},
                                     ignore_index=True)
            day_of_dm += 1

    res_df.to_csv(prefix_path + '/Output/dm date dict2.csv')
    print 'Done'


if __name__ == '__main__':
    main()