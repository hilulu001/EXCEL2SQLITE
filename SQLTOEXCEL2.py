# -*- coding: utf-8 -*-
import os
import xlrd
import xlsxwriter

import sqlite3
import time
import configparser
import re
def cx_zb(totalCount):
    strx1 = re.sub("\d", "", totalCount)
    strx2 = int(re.sub("\D", "", totalCount))+1
    return  strx1+str(strx2)
def cx_zb1(totalCount):
    strx1 = re.sub("\d", "", totalCount)
    return  strx1+':'+strx1
def cx_zb2(totalCount):
    strx1 = re.sub("\d", "", totalCount)
    print(strx1)
    return  chr(ord(strx1)+1)
def conf(msini):
    confbj=[]
    conf = configparser.ConfigParser()
    conf.read(os.getcwd() + '\\db\\'+msini, encoding="utf-8")
    in_file = conf.get('filename', 'in')
    out_file = conf.get('filename', 'out')
    sfbj= conf.get('filename', '是否插入')
    confbj.append(in_file)
    confbj.append(out_file)
    confbj.append(sfbj)
    return confbj
def wf_maincx( msfile1):
    msfile = os.getcwd() + '\\db\\配置表1.xlsx'
    wb = xlrd.open_workbook(filename=msfile)
    sheet=wb.sheet_by_index(0)

    msfile1 = os.getcwd() + '\\out\\' + msfile1
    print(msfile1)
    workbook = xlsxwriter.Workbook(msfile1)

    for y in range(1, sheet.nrows):
        print(sheet.row_values(y))
        sla=sheet.row_values(y)
        shtname = sla[0]
        zbs = sla[1]
        sqlx = sla[2]
        stypes = sla[3]
        cur.execute(sqlx)
        rows = cur.fetchall()
        names = [description[0] for description in cur.description]
        print(names)

        worksheet = workbook.add_worksheet(shtname)
        cell_format = workbook.add_format({'border': 1, 'bold': True, 'align': 'center', 'font_color': 'red','bg_color':'cccccc','font_size':'14','font_name':'微软雅黑'})
        worksheet.write_row(zbs, names, cell_format)

        cell_format1 = workbook.add_format({'border': 1,  'num_format': '0.00','font_size':'14'})
        zbsx=cx_zb(zbs)
        for row in rows:
            worksheet.write_row(zbsx, row, cell_format1)
            zbsx = cx_zb(zbsx)


        cell_format2 = workbook.add_format({'num_format': '0.00%'})

        worksheet.set_column('B:B', 20,cell_format2 )
        worksheet.set_column('A:B', 20)

    workbook.close()


if __name__ == '__main__':
    start = time.time()
    msini = "config.ini"
    confh=conf(msini)
    #print(confh)
    msfile =confh[1]
    print('开始时间：' + time.strftime("%Y_%m-%d %H:%M:%S", time.localtime()))
    ldb = os.getcwd() + '\\db\\myanswer.db'
    conn = sqlite3.connect(ldb)
    cur = conn.cursor()
    wf_maincx(msfile)
    conn.close()
    end = time.time()
    print('完成时间：' + time.strftime("%Y_%m-%d %H:%M:%S", time.localtime()))
    lasttime = int((end - start))
    print('耗时' + str(lasttime) + '秒')






