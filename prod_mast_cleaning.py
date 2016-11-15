#!/usr/bin/python
# -*- coding: gbk -*-

import xlrd
import sys
import csv

prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def load_sheet(path):
    workbook = xlrd.open_workbook(path)
    mast_file = workbook.sheet_by_name('Master file20140422')
    return mast_file


# no use
def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    path = prefix_path + '/Raw Data/Copy of Product Master-20160722(for RC).xlsx'

    mast_file = load_sheet(path)

    # have bug in handling u'\ue407'
    with open(prefix_path + '/Master File/Product Master.csv', 'wb') as f:
        wr = csv.writer(f, quoting=csv.QUOTE_ALL)

        for row_num in range(0, mast_file.nrows):
            wr.writerow(mast_file.row_values(row_num))

    print 'Done'


if __name__ == '__main__':
    main()