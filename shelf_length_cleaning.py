#!/usr/bin/python
# -*- coding: gbk -*-
import pandas as pd
import glob


def batch_process_shelf_length(dir_path):

    all_data = pd.DataFrame()

    for f in glob.glob(dir_path):
        df = pd.read_csv(f, encoding='gbk')
        fill_blank(df)
        month = f.split()[-3].split("\\")[-1]
        df["Month"] = month
        # print(month)
        # find data missing stores
        find_data_missing_store(df, month)
        all_data = all_data.append(df, ignore_index=True)

    all_data.to_csv("C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/Shelf Length/P2Y Shelf Length.csv", encoding='gbk')


def fill_blank(df):
    df["SFA Code"].fillna(method='pad', inplace=True)
    df["Store Name"].fillna(method='pad', inplace=True)


def find_data_missing_store(df, month):
    missing_list = df[(df["Category"] == u"总数") & (df["P&G Shelf Length"] == 0) & (df["All Shelf Length"] == 0)][
        ["Store Name", "SFA Code"]]
    missing_num = missing_list["SFA Code"].count()
    store_num = len(df["SFA Code"].unique())
    t_path = "C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/Shelf Length/Time %s Total %d Missing %d.csv" % (month, store_num, missing_num)
    missing_list.to_csv(t_path, encoding='gbk')


def main():
    # strict dir_path
    dir_path = "C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/Shelf Length/201*.csv"
    # folder_path = "C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/Display Num/"
    # test_path = "C:/Users/liu.y.25/Desktop/Promo Forecast/2 DM Data/Fem/Display Num/201501 Display Num.csv"
    batch_process_shelf_length(dir_path)

if __name__ == '__main__':
    main()
