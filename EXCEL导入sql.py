# -*- coding: utf-8 -*-
import sqlite3
import time
import os
import xlrd

import configparser
def conf(msini):
    confbj=[]
    conf = configparser.ConfigParser()
    conf.read(os.getcwd() + '\\db\\'+msini, encoding="utf-8")
    in_file = conf.get('filename', 'in')
    out_file = conf.get('filename', 'in')
    sfbj= conf.get('filename', '是否插入')
    confbj.append(in_file)
    confbj.append(out_file)
    confbj.append(sfbj)
    return confbj


def wf_main(msfile):
    mssheets = []
    msfile = os.getcwd() + '\\in\\' + msfile
    if not os.path.exists(msfile):
        print("★★★★★★不存在" + msfile)
        time.sleep(15)
        return
    wb = xlrd.open_workbook(filename=msfile)
    # 跑表
    mssheets = wb.sheet_names()
    for mssheet in mssheets:
        msfileds = []       # 跑表
        print(mssheet)
        sheet = wb.sheet_by_name(mssheet)
        # 跑字段，建表
        if sheet.ncols > 1:
            sql1 = ''
            sql2 = ''
            bl = []
            for x in range(0, sheet.ncols):
                print(sheet.cell(0, x).value)
                msfileds.append(sheet.cell(0, x).value + str(x))
                bl.append("?")
            # print(msfileds)
            bl = ','.join(bl)
            msfileds = ','.join(msfileds)
            sql1 = "CREATE TABLE IF NOT EXISTS  " + mssheet + " ( " + msfileds + " ); "
            sql2 = "INSERT INTO " + mssheet + " VALUES ( " + bl + " )"
            print(sql1)
            cur.execute(sql1)
            for y in range(1, sheet.nrows):
                cur.execute(sql2, tuple(sheet.row_values(y)))
                print(sheet.row_values(y))
            conn.commit()

if __name__ == '__main__':
    start = time.time()

    msini = "config.ini"
    confh=conf(msini)
    print(confh)
    msfile =confh[0]
    if  confh[2]=='否':
        if os.path.exists(os.getcwd() + '\\db\\myanswer.db'):
            os.remove(os.getcwd() + '\\db\\myanswer.db')

    print('开始时间：' + time.strftime("%Y_%m-%d %H:%M:%S", time.localtime()))
    ldb = os.getcwd() + '\\db\\myanswer.db'
    conn = sqlite3.connect(ldb)
    cur = conn.cursor()
    wf_main(msfile)
    conn.close()
    end = time.time()
    print('完成时间：' + time.strftime("%Y_%m-%d %H:%M:%S", time.localtime()))
    lasttime = int((end - start))
    print('耗时' + str(lasttime) + '秒')






