#!/usr/bin/python
# -*- coding: gbk -*-

import sys
import pandas as pd


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
dm_folder = prefix_path + '/DM POS Mapping excel(sihua)/'
edlp_folder = prefix_path + '/EDLP POS Mapping excel(sihua)/'

cols_name = ['srce_sys_id',
             'cust_extrn_id',
             'prod_extrn_id',
             'sales_date',
             'display_type']
srce_sys_id = '1124'