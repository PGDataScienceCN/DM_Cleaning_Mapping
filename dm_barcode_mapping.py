#!/usr/bin/python
# -*- coding: gbk -*-

import sys
import pandas as pd
from datetime import datetime, date, timedelta
import xlrd
import glob
import math


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    bigu_prod_mast = pd.read_csv(prefix_path + '/Master File WM/WM Product Master in BigU Y1516.csv')

    dm15 = pd.read_csv(prefix_path + '/Output/WM# Tab Performance -2015.csv')
    dm16 = pd.read_csv(prefix_path + '/Output/WM# Tab Performance -2016.csv')

    unmatched = pd.DataFrame(columns=['Year', 'Barcode'])
    multi_matched = pd.DataFrame(columns=['Year', 'Barcode', 'prod_extrn_id'])

    for barcode in dm15['Item barcode'].unique():
        if math.isnan(barcode):
            continue

        if barcode not in bigu_prod_mast['gtin']:
            unmatched = unmatched.append({'Year': '2015',
                                          'Barcode': int(barcode)},
                                         ignore_index=True)
        elif bigu_prod_mast[bigu_prod_mast['gtin'] == barcode]['prod_extrn_id'].__len__() > 1:
            for prod_extrn_id in bigu_prod_mast[bigu_prod_mast['gtin'] == barcode]['prod_extrn_id']:
                multi_matched = multi_matched.append({'Year': '2015',
                                                      'Barcode': barcode,
                                                      'prod_extrn_id': prod_extrn_id},
                                                     ignore_index=True)

    for barcode in dm16[' Item barcode '].unique():
        if math.isnan(barcode):
            continue
        if barcode not in bigu_prod_mast['gtin']:
            unmatched = unmatched.append({'Year': '2016',
                                          'Barcode': int(barcode)},
                                         ignore_index=True)
        elif bigu_prod_mast[bigu_prod_mast['gtin'] == barcode]['prod_extrn_id'].__len__() > 1:
            for prod_extrn_id in bigu_prod_mast[bigu_prod_mast['gtin'] == barcode]['prod_extrn_id']:
                multi_matched = multi_matched.append({'Year': '2016',
                                                      'Barcode': barcode,
                                                      'prod_extrn_id': prod_extrn_id},
                                                     ignore_index=True)

    unmatched.to_csv(prefix_path + '/Output/Unmatched SKU ID.csv')
    multi_matched.to_csv(prefix_path + '/Output/Multi Matched SKU ID.csv')

    print 'Done'

if __name__ == '__main__':
    main()