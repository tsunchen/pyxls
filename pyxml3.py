#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: tsunc & zadmine
@software: PyCharm Community Edition
@time: 2018/3/23 12:10
"""

import numpy as np
import pandas as pd
from lxml import objectify

# 从xml中提取有效数据，放入数组xmlist缓存
xmlist = []

filename = '8.xml'

srcfile = './test/' + filename

xml = objectify.parse(srcfile)
root = xml.getroot()
for element in root.iter():
     xmlist.append(element.text)
     #print("%s >>> %s" % (element.tag, element.text))

# 数组定位重新设计
bix = xmlist.index('00:00:00') - 10
bix

xmlist.reverse()
#xmlist

#len(xmlist)
eix = xmlist.index('00:00:00') - 17
eix

xmlist.reverse()
#xmlist
bix
eix 
len(xmlist)


# 中间值列值名称
fcnt = 27
cols = ['f1','f2','f3','f4','f5','f6','f7','f8','f9','f10','f11','f12','f13','f14','f15','f16','f17','f18','f19','f20','f21','f22','f23','f24','f25','f26','f27']

df = pd.DataFrame()

bbix = int(bix)
eeix = len(xmlist) - eix - 1

a = np.array(xmlist[bbix:eeix])
icnt = int(len(xmlist[bbix:eeix])/fcnt)
# 以i，fcnt为行列，重塑矩阵
b = a.reshape(icnt, fcnt)
#b

sess = []
for i in range(len(b)):
    #print (b[i])
    k = b[i]
    #k = calMega(b[i])
    #print (k)
    serr = pd.Series(k);
    sess.append(serr)

df = pd.concat(sess, axis=1)
df
df.index = cols
dt = df.T

idt = pd.DataFrame(dt, columns=['f3', 'f17', 'f19', 'f21', 'f23', 'f25', 'f27'])
idt

# 处理M，K，B数据，返回值
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


# 运算有效数据矩阵存放
res = []
valid_cols = 7

for row_c in range(icnt):
    for col_c in range(valid_cols):
        if ('-' not in idt.values[row_c][col_c]):
            #print (calMega(idt_2.values[row_c][col_c]))
            res.append(str(calMega(idt.values[row_c][col_c])))
        else:
            #print (idt_2.values[row_c][col_c])
            res.append(str(idt.values[row_c][col_c]))


# 矩阵转换
o = np.array(res)
out = o.reshape(icnt, valid_cols)

len(o)
len(out)


# 结果输出
sess_3 = []

for j in range(len(out)):
    #print (k)
    m = out[j]
    serr = pd.Series(m);
    sess_3.append(serr)

df_3 = pd.concat(sess_3, axis=1)
df_3

dt_3 = df_3.T
dt_3.columns.values

#dt_3.rename(columns={'A': 0, 'B': 1, 'C':2 }, inplace = True)
dt_3.columns = ['UserID','LastInbound','AvgInbound','MaxInbound','LastOutbound','AvgOutbound','MaxOutbound']
dt_3
dt_3.to_excel("xmltest.xlsx", filename)
