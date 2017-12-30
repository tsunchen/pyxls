#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: tsunz
@software: PyCharm Community Edition
@time: 2017/9/30 18:45

2017-12-30 update, process 14 data [for r in range(20, 34):]
                 , calMega [if (d != '0'):] [return ("0")]

"""

import xlrd,xlwt

def calMega(d):
   if (d != '0'):
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
   return ("0")


#17 - 31

def pretotalnoshare(table):
    for r in range(20, 34):
         print (table.col_values(0)[r])
         print (':  Inbound Avg:', table.col_values(8)[r], '->' , calMega(table.col_values(8)[r]))
         #print (calMega(table.col_values(8)[r]))
         print (':  Inbound Max:', table.col_values(9)[r], '->' , calMega(table.col_values(9)[r]))
         #print (calMega(table.col_values(9)[r]))

def savpretotalnosharebound(ff, table):
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('NoshareBound-Total')
    for r in range(20, 34):
         #print (table.col_values(0)[r])
         L1 = table.col_values(0)[r]
         L2 = table.col_values(9)[r]
         L3 = table.col_values(8)[r]
         L4 = calMega(L2)
         L5 = calMega(L3)
         #print (':  Inbound Avg:', table.col_values(8)[r], '->' , calMega(table.col_values(8)[r]))
         #print (calMega(table.col_values(8)[r]))
         #print (':  Inbound Max:', table.col_values(9)[r], '->' , calMega(table.col_values(9)[r]))
         #print (calMega(table.col_values(9)[r]))
         i = 0
         worksheet.write(r - 20, i, label=L1)
         i = i + 1
         worksheet.write(r - 20, i, label=L2)
         i = i + 1
         worksheet.write(r - 20, i, label=L3)
         i = i + 1
         worksheet.write(r - 20, i, label=L4)
         i = i + 1
         worksheet.write(r - 20, i, label=L5)
    workbook.save(ff + 'Result-NBT.xls')



if __name__ == '__main__':
    file_folder = r"E:\backupswitch\localroot\pub\easybridge-BA\nosharebound-total\2017-Q4" + '\\'
    file_name = r'Cacti_report_2017_11P.xlsx'

    infile = file_folder + file_name
    data = xlrd.open_workbook(infile)
    table = data.sheets()[0]
    pretotalnoshare(table)
    savpretotalnosharebound(file_folder, table)
