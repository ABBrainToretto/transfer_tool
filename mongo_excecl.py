# _*_ coding:utf-8 _*_

import xlwt
from pymongo import MongoClient

def get_connect():
    client = MongoClient()
    db = client['taobao']
    collention = db['sijin']

def create_xls():
    heads = ['shop','deal','title','price','id','location','image']
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet(u'sijin',cell_overwrite_ok=True)

    # 建立表头字段
    i = 0
    for head in heads:
        sheet.write(0,i,head)
        i += 1
    contents = get_connect().collention.find()
    # for content in contents:
    #     print content['title']

    # 插入数据
    t = 1
    for content in contents:
        sheet.write(t,0,content['shop'])
        sheet.write(t,1,content['deal'])
        sheet.write(t,2,content['title'])
        sheet.write(t,3,content['price'])
        sheet.write(t,4,content['id'])
        sheet.write(t,5,content['location'])
        sheet.write(t,6,content['image'])
        t += 1

    # 保存在excel中
    wbk.save('taobao_sijin.xls')

create_xls()
