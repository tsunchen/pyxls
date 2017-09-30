#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: tsunz
@software: PyCharm Community Edition
@time: 2017/9/30 18:45

"""

import xlrd

def calMega(d):
   if 'M' in d:
     #print ('M')
     return (d[:-2])
   elif 'k' in d:
     #print ('k')
     #print (d[:-2])
     kilobyte = float(d[:-2]) / 1000
     return (kilobyte)
   else:
     #print (d)
     #print ('no M or k')
     byte = float(d[:-2]) / 1000
     byte2 = byte / 1000
     return ('%.9f' % byte2)



#17 - 31

def pretotalnoshare(table):
    for r in range(17, 31):
         print (table.col_values(0)[r])
         print (':  Inbound Avg:', table.col_values(8)[r], '->' , calMega(table.col_values(8)[r]))
         #print (calMega(table.col_values(8)[r]))
         print (':  Inbound Max:', table.col_values(9)[r], '->' , calMega(table.col_values(9)[r]))
         #print (calMega(table.col_values(9)[r]))


if __name__ == '__main__':
    file_name = r'cacti_report_13-201707.xlsx'
    file = r"E:\backupswitch\localroot\pub\easybridge-BA\noshare-total\2017-07,08" + '\\' + file_name
    data = xlrd.open_workbook(file)
    table = data.sheets()[0]
    pretotalnoshare(table)
