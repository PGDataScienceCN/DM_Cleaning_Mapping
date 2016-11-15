#!/usr/bin/python
# -*- coding: gbk -*-

import xlrd
import sys
from datetime import date, datetime
import csv
from NO_USE_dm_fmot_upc_mapping import timestamp_to_date


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'


def load_sheet(path):
    workbook = xlrd.open_workbook(path)
    raw_sheet = workbook.sheet_by_name('raw data')
    return workbook, raw_sheet


def get_date(raw_sheet, workbook, row_num, col_num):
    year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(int(raw_sheet.row_values(row_num)[col_num]), workbook.datemode)
    return date(year, month, day)


# raw data have exception
def get_date2(cell_value):
    print cell_value
    year = int(cell_value[0:4])
    month = int(cell_value[4:6])
    day = int(cell_value[6:8])
    print year, month, day
    return date(year, month, day)


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    path = prefix_path + '/Raw Data/SWF_EDLP_PDQ SKU.xlsx'

    workbook, raw_sheet = load_sheet(path)

    with open(prefix_path + '/DM Dictionary/WM Y1516 Long Term DM Dictionary.csv', 'wb') as f:
        wr = csv.writer(f, quoting=csv.QUOTE_ALL)
        wr.writerow(raw_sheet.row_values(0))

        for row_num in range(1, raw_sheet.nrows):
            # row_value = raw_sheet.row_values(row_num)
            # row_value[10:11] = get_date(raw_sheet, workbook, row_num, 10), get_date(raw_sheet, workbook, row_num, 11)

            # print str(raw_sheet.cell_value(row_num, 5))
            # have exception like 20160631
            # row_value[5:6] = get_date2(str(raw_sheet.cell_value(row_num, 5))), get_date2(str(raw_sheet.cell_value(row_num, 6)))

            wr.writerow(row_value)

    print 'Done'

if __name__ == '__main__':
    main()