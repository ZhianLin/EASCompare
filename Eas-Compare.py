# -*- coding: utf-8 -*-
"""
Created on Fri Apr 21 13:27:36 2017

@author: Administrator
Eas配置批量比较
"""
"""
ToDoList:
1、	列出目录下所有的excel文件名
2、	读取一个excel文件中sheet名为“核算项目(非保证金类)”的整个表。（18列数据）
3、	把数据保存到数据库中（20列，增加文件名、地区）
4、	读取第二个excel文件的
5、	拼接成大的eas配置表A1
6、	读取系统数据库中的eas配置表A2
7、	按收费项目比较A1跟A2，相差一个单元格的都做为差异。

"""
import os,xlrd
#import xlwt
import glob
import re
import csv
import sys
'''
1、	列出目录下所有的excel文件名
'''
path0=u'E:/项目/珠江投资/交接文档/二期/金蝶税控开票/20161214/2016年11月17日极致对接科目变动/2016年11月17日极致对接科目变动/广州/'
path1=u'E:/项目/珠江投资/交接文档/二期/金蝶税控开票/20161214/2016年11月17日极致对接科目变动/2016年11月17日极致对接科目变动/北京/'
path2=u'F:/2016年11月17日极致对接科目变动/上海/'
path1=u'G:/极致软件/项目/珠江投资/交接文档/二期/金蝶税控开票/20170427华北华东的电费计划收费项目变更/2017年4月27日极致对接科目变动/广州/'
filelist=[]
for root,dirs,files in os.walk(path1):
    for filename in files:
        filelist.append(path1+filename)
#print(filelist)
#print(len(filelist))
'''
2、	读取一个excel文件中sheet名为“核算项目(非保证金类)”的整个表。（18列数据）
3、	把数据保存到数据库中（20列，增加文件名、地区）
'''
filedata=[]
for i in range(len(filelist)):
    #print (filelist[i])    
    wb=xlrd.open_workbook(filelist[i])
    #print('wb',wb)
    sh=wb.sheet_by_name(u'核算项目(非保证金类)')
    for rownum in range(sh.nrows):
        #print(sh.row_values(rownum))
        #sh.cell_value()
        rowdata=sh.row_values(rownum)
        rowdata.insert(0,filelist[i].split('/')[-1])
        #print(rowdata)
        filedata.append(rowdata)
    
    #print('内大小',len(filedata))
    #'''

print('外大小',len(filedata))
r1=re.compile(r'.*(电费)')
r2=re.compile(r'.*(其他业务收入)')
dd=[]
for i in range(len(filedata)):
    d=filedata[i][3]
    d2=filedata[i][18]
    csvFile=open('e:/Eas201704211705.csv','wb')       
    write=csv.writer(csvFile,dialect='excel',quotechar='|',quoting=csv.QUOTE_MINIMAL
                     )
    print('文件名:%s，\n收费项目：%s，计划流入项目：%s\n'%(filedata[i][0].split('/')[-1],d,d2))
    if r1.findall(d) :
        print('匹配到公摊')
        #print(d2)
        if r2.findall(d2):
            print(filedata[i][0],'匹配到公摊',filedata[i][18])
            print(type(filedata[i]))
            write.writerow(('a'))            
        elif r2.findall(filedata[i][17]):
            print(filedata[i][0],'匹配到公摊',filedata[i][17])
        else:print('匹配不到“其他业务收入”')
    #else:print('%s匹配不到公摊'%filedata[i][0])    
    print('==========')
    csvFile.close()
    
#print('0\n',filedata[0][1])    
#print('1\n',filedata[1][2])  
#print('2\n',filedata[2][2])    
#print('3\n',filedata[3][2])  
