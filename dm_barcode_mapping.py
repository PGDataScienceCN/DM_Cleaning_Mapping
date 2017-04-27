#!/usr/bin/python
# -*- coding: gbk -*-

import sys
import pandas as pd
from datetime import datetime, date, timedelta
import xlrd
import glob
import math
import numpy


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    bigu_prod_mast = pd.read_csv(prefix_path + '/Master File WM/WM Product Master in BigU Y1516.csv')

    dm15 = pd.read_csv(prefix_path + '/Output/WM# Tab Performance -2015.csv')
    dm16 = pd.read_csv(prefix_path + '/Output/WM# Tab Performance -2016.csv')

    unmatched = pd.DataFrame(columns=['Year', 'UPC', 'Barcode', 'Item Desc'])
    multi_matched = pd.DataFrame(columns=['Year', 'UPC', 'Barcode', 'prod_extrn_id'])



    for upc in dm15[' UPC '].unique():
        if isinstance(upc, str):
            continue
        if math.isnan(upc):
            continue

        if int(upc) not in bigu_prod_mast['bar_code']:
            unmatched = unmatched.append({'Year': '2015',
                                          'UPC': upc,
                                          'Barcode': dm15[dm15[' UPC '] == upc]['Item barcode'],
                                         'Item Desc': dm15[dm15[' UPC '] == upc][' Promotion Description ']},
                                         ignore_index=True)
        elif bigu_prod_mast[bigu_prod_mast['bar_code'] == upc]['prod_extrn_id'].__len__() > 1:
            for prod_extrn_id in bigu_prod_mast[bigu_prod_mast['gtin'] == upc]['prod_extrn_id']:
                multi_matched = multi_matched.append({'Year': '2015',
                                                      'UPC': upc,
                                                      'Barcode': dm15[dm15[' UPC '] == upc]['Item barcode'],
                                                      'prod_extrn_id': prod_extrn_id},
                                                     ignore_index=True)

    for upc in dm16[' UPC '].unique():

        if math.isnan(float(upc)):
            continue

        if int(upc) not in bigu_prod_mast['bar_code'].tolist():
            unmatched = unmatched.append({'Year': '2016',
                                          'UPC': upc,
                                          'Barcode': dm16[dm16[' UPC '] == upc][' Item barcode '],
                                          'Item Desc': dm16[dm16[' UPC '] == upc][' Promotion Description ']},
                                         ignore_index=True)
        elif bigu_prod_mast[bigu_prod_mast['bar_code'] == upc]['prod_extrn_id'].__len__() > 1:
            for prod_extrn_id in bigu_prod_mast[bigu_prod_mast['gtin'] == upc]['prod_extrn_id']:
                multi_matched = multi_matched.append({'Year': '2016',
                                                      'UPC': upc,
                                                      'Barcode': dm16[dm16[' UPC '] == upc][' Item barcode '],
                                                      'prod_extrn_id': prod_extrn_id},
                                                     ignore_index=True)

    unmatched.to_csv(prefix_path + '/Output/Unmatched SKU ID2.csv', encoding="gbk")
    multi_matched.to_csv(prefix_path + '/Output/Multi Matched SKU ID2.csv', encoding="gbk")

    print 'Done'

if __name__ == '__main__':
    main()