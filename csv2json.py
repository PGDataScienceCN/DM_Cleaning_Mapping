#!/usr/bin/python
# -*- coding: gbk -*-

import csv
import json
import sys

prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/'


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    csvfile = open(prefix_path + 'Knime Opt/dm sku profile2.csv', 'r')
    jsonfile = open(prefix_path + 'Knime Opt/dm_sku_profile.json', 'w')

    fieldnames = ("promo_desc", "upc", "A1", "A2", "gtin", "prod_name", "usage", "report_size", "sku#")
    reader = csv.DictReader(csvfile, fieldnames)
    for row in reader:
        json.dump(row, jsonfile)
        jsonfile.write('\n')

    print 'Done'

if __name__ == '__main__':
    main()