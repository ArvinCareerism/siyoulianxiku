# -*- coding: utf-8 -*-
import xlrd
import xlwt
from datetime import date,datetime
from xlutils.copy import copy
import os

path1 = (r'E:\lianxi\PY\五金生产日报.xls')
rbook = xlrd.open_workbook(path1)
#print (rbook.sheet_names())
path2 = (r'E:\lianxi\PY\五金生产效率考核数据记录表.xls')
zbook = xlrd.open_workbook(path2,formatting_info = True )
wbook = copy(zbook)

sheet1 = rbook.sheet_by_index(2)

#rows = sheet1.row_values(3) # 获取第四行内容
cols = sheet1.col_values(16) # 获取第三列内容
index1 = len(cols)
#print(index1)
#print ("******")
datas = []
for i in range(2,index1):
    #print(sheet1.row_values(i))
    datas.append(sheet1.row_values(i))
'''
print(datas[8][16])
print(datas[8][5])
print(datas[8][9])
print(datas[8][10])
print ("******")
print(cols[2])
print(cols[3])
print(cols[4])
print(bool(cols[2]!=cols[3]))
#print (rows)
#print (cols)
#print (type(datas))
#print (datas)
'''


wsheet_name = wbook.get_sheet(2)
#wsheet_name.write_formula(3,2,'=B3 + B4')
tt = '=L7 + M7\n'
wsheet_name.write(3,2,tt)
r = 0
n = 10
index2 = 0
while index2 < index1:
    #print ("################")
    print(index2,r,n)

    index2 += 1

    if cols[index2] == cols[index2+1]:
        n += 4
        #print("&&&&&")
        #print (index2,r,n)
        wsheet_name.write(r+2,n,datas[index2-1][5])
        wsheet_name.write(r+2,n+1,datas[index2-1][9])
        wsheet_name.write(r+2,n+2,datas[index2-1][10])
    else:
        r += 1
        wsheet_name.write(r+2,1,datas[index2-1][16])
        wsheet_name.write(r+2,10,datas[index2-1][5])
        wsheet_name.write(r+2,11,datas[index2-1][9])
        wsheet_name.write(r+2,12,datas[index2-1][10])

        n = 10
        #print ("**********")
        #print (index2,r,n)
    if len(datas[index2][16])==0:
        break


os.remove(path2)
wbook.save(path2)

