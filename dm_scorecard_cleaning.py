#!/usr/bin/python
# -*- coding: gbk -*-

import xlrd
import sys
from datetime import date, datetime
import csv
import glob

prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def load_sheet(path):
    workbook = xlrd.open_workbook(path)
    raw_sheet = workbook.sheet_by_name('问卷明细')
    return raw_sheet


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    path = prefix_path + '/Raw Data/DM1516/WM 1*.xlsx'
    dm_num = 0
    add_dm_flag = False

    for f in glob.glob(path):
        # get dm number
        if f.split()[-3] == '2':
            dm_num = f.split()[-4]
            add_dm_flag = True
            month1 = f.split()[-1].split('.')[0]
            month2 = f.split()[-1].split('-')[1].split('.')[0]

        else:
            dm_num = f.split()[-2][0:4]
            add_dm_flag = False
            month1 = f.split()[-1].split('.')[0]
            month2 = f.split()[-1].split('-')[1].split('.')[0]

        raw_sheet = load_sheet(f)
        # (2,0) is sfa code
        # (3,_) is dm item
        # (4,2) is project task
        cols_name = list(['SFA Code', 'Store Name', 'Project Task']) + raw_sheet.row_values(3)[3:]

        with open(prefix_path + '/DM Scorecard/WM %s %d DM Scorecard %s %s.csv' % (dm_num, add_dm_flag, month1, month2), 'wb') as dm:
            wr = csv.writer(dm, quoting=csv.QUOTE_ALL)
            wr.writerow(cols_name)
            for row_num in range(5, raw_sheet.nrows):
                # print row_num
                wr.writerow(raw_sheet.row_values(row_num))

        print dm_num

    print 'Done'

if __name__ == '__main__':
    main()