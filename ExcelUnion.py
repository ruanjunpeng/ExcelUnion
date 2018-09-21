# -*- coding: utf-8 -*-
"""
Created on Fri Sep 21 21:22:44 2018

@author: JunpengRuan
"""

'''
表格合并辅助小工具，输入整数N，将所有表格的第N行复制到第一个表格的下面
'''
import os
import sys
import xlwt
import xlrd

Dir = sys.argv[1]
N = sys.argv[2]
N = int(N)
filenames = list()
for file in os.listdir(Dir):
    r=file.decode('GB2312')
    filenames.append(r)
    #上面的不知道为啥用CMD没有问题但是用IPython console乱码
for file in filenames:
    print file

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('Sheet1')

insertRowNo = 0
for file in filenames:
    print Dir +'\\'+ file
    wb = xlrd.open_workbook(Dir+'\\'+file)
    ws = wb.sheet_by_index(0)
    data = ws.row_values(N-1)
    for i in range(len(data)):
        if(type(data[i]).__name__ == 'float'):
            data[i]=unicode(int(data[i]))
    for cell in data:
        print type(cell)
    for i in range(len(data)):
        worksheet.write(insertRowNo,i,data[i])
    insertRowNo = insertRowNo + 1
workbook.save(Dir + '\\' + 'AllInOne.xls')