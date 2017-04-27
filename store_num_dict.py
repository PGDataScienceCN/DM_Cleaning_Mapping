#!/usr/bin/python
# -*- coding: gbk -*-

import sys
import pandas as pd
from datetime import datetime, date, timedelta
import xlrd
import glob


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
cols_name = ['store_num',
             'cust_extrn_id']

def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    bigu_mast_store_df = pd.read_csv(prefix_path + '/Master File WM/WM Store Master in BigU.csv')
    wm_store_mast_sheet = xlrd.open_workbook(
        prefix_path + '/Raw Data/Copy of 11月3日 Store Profile new format v2.xlsx').sheet_by_name(u'SC ')

    t_df = pd.DataFrame(columns=cols_name)

    for row in range(1, wm_store_mast_sheet.nrows):
        store_num = int(wm_store_mast_sheet.row(row)[0].value)

        cust_extrn_id = bigu_mast_store_df[bigu_mast_store_df['cust_extrn_dads.cust_store_num'] == store_num]['cust_extrn_dads.cust_extrn_id'].tolist()[0]
        print cust_extrn_id
        t_df = t_df.append({'store_num': str(store_num),
                            'cust_extrn_id': str(cust_extrn_id)},
                           ignore_index=True)
        t_df.to_csv(prefix_path + '/Output/DM FMOT/Store Num Dict.csv')

    print 'Done'


if __name__ == '__main__':
    main()