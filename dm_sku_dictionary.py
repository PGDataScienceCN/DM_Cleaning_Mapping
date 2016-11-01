#!/usr/bin/python
# -*- coding: gbk -*-
import pandas as pd
import csv
import numpy
import glob


def get_dm_sku_dictionary(dir_path):
    all_data = pd.DataFrame()
    lst = list()

    for f in glob.glob(dir_path):
        period = f.split()[-1].split('-')
        time = f.split()[-2]
        start_date = period[0]
        end_date = period[1].replace('.csv', '')

        df = pd.read_csv(f, encoding='gbk')
        lst += df.columns.tolist()
        break

    # remove unnecessary items
    new_lst = [i for i in lst if 'Unnamed:' not in i]
    new_lst.remove('SFA Code')
    new_lst.remove('Store Name')
    new_lst.remove('Project Task')

    # TODO: suspend to wait for UPC mapping
    print (len(new_lst))
    with open("test.csv", 'w') as f:
        for i in new_lst:
            f.write(('%s\n' % i).encode('gbk'))


def main():
    # strict dir_path
    # t_path = "C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/DM Scorecard/WM 1614 06.23-07.06.csv"
    dir_path = "C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/DM Scorecard/WM 1*.csv"
    get_dm_sku_dictionary(dir_path)

if __name__ == '__main__':
    main()