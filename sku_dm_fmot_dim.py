#!/usr/bin/python
# -*- coding: gbk -*-

import sys
import pandas as pd
import glob


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'

cols_name = ['srce_sys_id',
             'cust_extrn_id',
             'prod_extrn_id',
             'bigu_gtin',      # map with barcode, I don't know why :(
             'bigu_barcode',   # map with upc, I don't know why :(
             'bigu_item_desc',
             'upc',
             'barcode',
             'item_desc',
             'dm_num',
             'sales_date',
             'dm_question',
             'display_type']


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    res_df = pd.DataFrame(columns=cols_name)

    for f in glob.glob(prefix_path + '/Output/DM FMOT/SKU DM 1*.csv'):
        res_df = res_df.append(pd.read_csv(f, dtype={'cust_extrn_id': object, 'prod_extrn_id': object, 'bigu_gtin': object, 'bigu_barcode': object, 'upc': object, 'barcode': object}))

    res_df.to_csv(prefix_path + '/Output/SKU DM Dict.csv')
    # res_df.to_excel(prefix_path + '/Output/SKU DM Dict.xlsx')

    print 'Done'


if __name__ == '__main__':
    main()