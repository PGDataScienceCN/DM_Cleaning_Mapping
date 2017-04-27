#!/usr/bin/python
# -*- coding: gbk -*-

import sys
import pandas as pd
from datetime import datetime, date, timedelta
import xlrd
import glob


prefix_path = 'C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem'
dm_folder = prefix_path + '/DM POS Mapping excel(sihua)/'
edlp_folder = prefix_path + '/EDLP POS Mapping excel(sihua)/'

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
srce_sys_id = '1124'


def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + timedelta(n)


def pampers_reformat(org_str):
    nick_name1 = u'帮宝适特级棉柔（白帮）'
    nick_name2 = u'帮宝适超薄干爽(绿帮)'
    if nick_name1 in org_str:
        return org_str.replace(nick_name1, u'白帮')
    elif nick_name2 in org_str:
        return org_str.replace(nick_name2, u'绿帮')
    else:
        return org_str


def load_dm_mapping(dm_num):
    f_dir = dm_folder + 'DM %s UPC Dictionary*.xlsx' % dm_num
    for path in glob.glob(f_dir):
        t_df = pd.read_excel(path, sheetname='Sheet1', header=None, names=['Question', 'Item Desc', 'UPC', 'Barcode'])
        return t_df


def get_cust_extrn_id(st, bigu_mast_store_df, wm_store_mast_sheet):
    # modify wrong sfa code
    if int(st) == 4432:
        st = 35190
        print '4432 has been to 35190'

    for row in range(wm_store_mast_sheet.nrows):
        row_value = wm_store_mast_sheet.row(row)
        if row_value[1].value == int(st):
            store_num = row_value[0].value

    t_cust_extrn_id = bigu_mast_store_df[bigu_mast_store_df['cust_extrn_dads.cust_store_num'] == store_num]['cust_extrn_dads.cust_extrn_id'].tolist()

    try:
        cust_extrn_id = t_cust_extrn_id[0]
    except IndexError:
        print int(st)
        return None, int(st), None

    return cust_extrn_id


def gen_fmot_sub_df(dm_num, start_date, end_date, cust_extrn_id, prod_extrn_id, bigu_gtin, bigu_barcode, bigu_item_desc, upc, barcode, item_desc, dm_question, display_type):
    t_df = pd.DataFrame(columns=cols_name)

    for sales_date in daterange(start_date, end_date):
        t_df = t_df.append({'srce_sys_id': srce_sys_id,
                            'cust_extrn_id': cust_extrn_id,
                            'prod_extrn_id': prod_extrn_id,
                            'bigu_gtin': bigu_gtin,
                            'bigu_barcode': bigu_barcode,
                            'bigu_item_desc': None,  # bigu_item_desc,  endoc has issue with bigu item desc
                            'upc': str(upc),
                            'barcode': str(barcode),
                            'item_desc': item_desc,
                            'dm_num': dm_num,
                            'sales_date': sales_date,
                            'dm_question': dm_question,
                            'display_type': display_type},
                           ignore_index=True)
    return t_df


def load_dm_csv(dm_num):
    folder_path = prefix_path + '/Raw Data/DM1516/WM %s *.xlsx' % dm_num
    for f in glob.glob(folder_path):
        workbook = xlrd.open_workbook(f)
        return workbook.sheet_by_name(u'问卷明细')


# TODO: 2 1561 dm csv
def load_1561_csv(dm_num):
    folder_path = prefix_path + '/Raw Data/DM1516/WM %s 1*.xlsx' % dm_num
    for f in glob.glob(folder_path):
        workbook = xlrd.open_workbook(f)
        return workbook.sheet_by_name(u'问卷明细')


def gen_dm_sku_store_fmot(dm_num, start_date, end_date, bigu_prod_mast_sheet, bigu_mast_store_df, wm_store_mast_sheet):

    res_df = pd.DataFrame(columns=cols_name)
    wrong_store_df = pd.DataFrame(columns=['SFA Code', 'DM'])
    wrong_sku_df = pd.DataFrame(columns=['DM', 'UPC', 'Desc'])

    dm_mapping_df = load_dm_mapping(dm_num)

    if dm_num == 1561:
        dm_sheet = load_1561_csv(dm_num)
    else:
        dm_sheet = load_dm_csv(dm_num)

    for row_idx in range(5, dm_sheet.nrows - 1):

        sfa_code = dm_sheet.row(row_idx)[0].value

        cust_extrn_id = get_cust_extrn_id(sfa_code, bigu_mast_store_df, wm_store_mast_sheet)

        if not cust_extrn_id:
            wrong_store_df = wrong_store_df.append({'SFA Code': sfa_code.value,
                                                    'DM': dm_num},
                                                   ignore_index=True)

        for q_idx in range(3, dm_sheet.row(3).__len__(), 3):
            dm_question = pampers_reformat(dm_sheet.row(3)[q_idx].value)

            display_type = dm_sheet.row(row_idx)[q_idx].value

            if dm_question in dm_mapping_df['Question'].tolist():
                for idx, item in dm_mapping_df[dm_mapping_df['Question'] == dm_question].iterrows():
                    item_desc = dm_mapping_df.get_value(idx, 'Item Desc')
                    upc = dm_mapping_df.get_value(idx, 'UPC')
                    barcode = dm_mapping_df.get_value(idx, 'Barcode')  # 6.9e+12 cannot use

                    # map tab performance UPC with bigu barcode
                    # prod_extrn_id, bigu_gtin, bigu_barcode, bigu_item_desc = get_bigu_prod_dt(upc, bigu_prod_mast_sheet)
                    #
                    # if not prod_extrn_id:
                    #     wrong_sku_df = wrong_sku_df.append({'DM': dm_num,
                    #                                         'UPC': upc,
                    #                                         'Desc': item_desc},
                    #                                         ignore_index=True)

                    for row_idx2 in range(1, bigu_prod_mast_sheet.nrows):
                        row_value = bigu_prod_mast_sheet.row(row_idx2)
                        if int(row_value[5].value) == int(upc):
                            prod_extrn_id = row_value[0].value
                            bigu_gtin = row_value[1].value
                            bigu_barcode = row_value[5].value
                            bigu_item_desc = row_value[4].value

                            t_df = gen_fmot_sub_df(dm_num, start_date, end_date, cust_extrn_id, prod_extrn_id, bigu_gtin,
                                                   bigu_barcode, bigu_item_desc, upc,
                                                   barcode, item_desc, dm_question, display_type)

                            res_df = res_df.append(t_df)
        #break  # one loop test

    return res_df, wrong_store_df, wrong_sku_df


# not use
def get_bigu_prod_dt(upc, bigu_prod_mast_sheet):
    prod_extrn_id = list()
    bigu_gtin = None
    bigu_barcode = None
    bigu_item_desc = None

    for row_idx in range(1, bigu_prod_mast_sheet.nrows):
        row_value = bigu_prod_mast_sheet.row(row_idx)

        if int(row_value[5].value) == int(upc):
            # print int(row_value[5].value)
            prod_extrn_id.append(row_value[0].value)

            bigu_gtin = row_value[1].value
            bigu_barcode = row_value[5].value
            bigu_item_desc = row_value[4].value

    if prod_extrn_id:
        return ','.join(str(x) for x in prod_extrn_id), bigu_gtin, bigu_barcode, bigu_item_desc
    else:
        print '??? ' + str(upc)
        return None, None, None, None


def main():
    reload(sys)
    sys.setdefaultencoding('gbk')

    res_df = pd.DataFrame(columns=cols_name)
    wrong_store_df = pd.DataFrame(columns=['SFA Code', 'DM'])

    bigu_prod_dict = pd.read_csv(prefix_path + '/Output/WM Product Dict in BigU Y1516.csv')
    bigu_prod_mast_sheet = xlrd.open_workbook(prefix_path + '/Master File WM/WM Product Master in BigU.xlsx').sheet_by_name(u'Sheet')

    bigu_mast_store_df = pd.read_csv(prefix_path + '/Master File WM/WM Store Master in BigU.csv')

    wm_store_prof_df = pd.read_csv(prefix_path + '/Master File WM/WM Store Profile.csv')  # no use
    wm_store_mast_sheet = xlrd.open_workbook(prefix_path + '/Raw Data/Copy of 11月3日 Store Profile new format v2.xlsx').sheet_by_name(u'SC ')

    dm_date_dict = pd.read_csv(prefix_path + '/Output/DM Date Dict.csv')
    dm_date_dict['Start Date'] = pd.to_datetime(dm_date_dict['Start Date'], format='%Y/%m/%d')
    dm_date_dict['End Date'] = pd.to_datetime(dm_date_dict['End Date'], format='%Y/%m/%d')

    # for single_date in daterange(dm_date_dict['Start Date'][2].date(), dm_date_dict['End Date'][2].date()):
    #     print single_date

    for index, row in dm_date_dict.iterrows():
        # if row['DM'] == 1615:
        #     print 'Start'
        #     res_df, wrong_store_df, wrong_sku_df = gen_dm_sku_store_fmot(row['DM'], row['Start Date'].date(), row['End Date'].date(), bigu_prod_mast_sheet, bigu_mast_store_df, wm_store_mast_sheet)
        #     print row['DM']
        #     res_df.to_csv(prefix_path + '/Output/DM FMOT/SKU DM %s FMOT DIM.csv' % row['DM'])
        #     wrong_store_df.to_csv(prefix_path + '/Output/DM FMOT/DM %s Wrong Stores.csv' % row['DM'])
        #     wrong_sku_df.to_csv(prefix_path + '/Output/DM FMOT/DM %s Wrong SKUs.csv' % row['DM'])

        res_df, wrong_store_df, wrong_sku_df = gen_dm_sku_store_fmot(row['DM'], row['Start Date'].date(), row['End Date'].date(), bigu_prod_mast_sheet, bigu_mast_store_df, wm_store_mast_sheet)

        print row['DM']
        res_df.to_csv(prefix_path + '/Output/DM FMOT/SKU DM %s FMOT DIM.csv' % row['DM'])
        wrong_store_df.to_csv(prefix_path + '/Output/DM FMOT/DM %s Wrong Stores.csv' % row['DM'])
        wrong_sku_df.to_csv(prefix_path + '/Output/DM FMOT/DM %s Wrong SKUs.csv' % row['DM'])
        # break  # one loop test

    print 'Done'


if __name__ == '__main__':
    main()