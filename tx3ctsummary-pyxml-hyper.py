#!/usr/bin/env python
# -*- coding:utf-8 -*-

"""
@author: tsunc & zadmine
@software: PyCharm Community Edition
@time: 2018/10/30 00:10
"""

import numpy as np
import pandas as pd
from lxml import objectify

import os, sys


#filename = 'cacti_report_12.xml'
#srcfile = './test/' + filename

#
# 扫描源文件夹
mapperfile = []
mapperpath = './mapper/'
#mapperpath = 'E:/tsun/ct-summary-py/mapper'

for fpathe, dirs, fs in os.walk(mapperpath):
    for f in fs:
        ##print (os.path.join(fpathe,f))
        #print f
        mapperfile.append(f)
print(mapperfile)

df_mapper = pd.read_excel(str(mapperpath)+'/'+mapperfile[0], sheet_name = 'sn-mapper-name')

#

# 扫描源文件夹
srcfiles = []
srcpath = './test/'
#srcpath = 'E:/tsun/ct-summary-py/test'

for fpathe, dirs, fs in os.walk(srcpath):
    for f in fs:
        ##print (os.path.join(fpathe,f))
        #print f
        srcfiles.append(f)
#print(srcfiles)


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

# 处理对照mapper表的name索引到sn
def pyxmlm_ms2003(srcfile, mapper):
    # 从xml中提取有效数据，放入数组xmlist缓存
    # xmlist = []
    xml = objectify.parse(srcfile)
    root = xml.getroot()
    '''
    for element in root.iter():
        xmlist.append(element.text)
        #print("%s >>> %s" % (element.tag, element.text))
    '''
    xmlist = [ element.text for element in root.iter() ]
    # 数组定位重新设计
    bix = xmlist.index('00:00:00') - 10
    #bix
    xmlist.reverse()
    #xmlist
    #len(xmlist)
    eix = xmlist.index('00:00:00') - 17
    #eix
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

    # 运算有效数据矩阵存放
    #res = []
    valid_cols = 7

    '''
    for row_c in range(icnt):
        for col_c in range(valid_cols):
            if ('-' not in idt.values[row_c][col_c]):
                #print (calMega(idt_2.values[row_c][col_c]))
                #res.append(str(calMega(idt.values[row_c][col_c]) + str('M')))
                res.append(str(calMega(idt.values[row_c][col_c])) + str('M'))
            else:
                #print (idt_2.values[row_c][col_c])
                res.append(str(idt.values[row_c][col_c]))
    '''
    res = [ str(calMega(idt.values[row_c][col_c])) if ('-' not in idt.values[row_c][col_c]) else str(idt.values[row_c][col_c]) for row_c in range(icnt) for col_c in range(valid_cols) ]


    # 矩阵转换
    o = np.array(res)
    out = o.reshape(icnt, valid_cols)

    len(o)
    len(out)


    # 结果输出
    sess_3 = []

    # 索引值对应
    for j in range(len(out)):#print (k)
        for index, r in mapper.iterrows():
            if (r['name'] == out[j][0]):
               mappersn = np.append(r['sn'], out[j])
               #m = out[j]
               #print(out[j][0])
               serr = pd.Series(mappersn)
               sess_3.append(serr)

    df_3 = pd.concat(sess_3, axis=1)
    df_3

    dt_3 = df_3.T
    dt_3.columns.values

    #dt_3.rename(columns={'A': 0, 'B': 1, 'C':2 }, inplace = True)
    dt_3.columns = ['YourMapperSN', 'UserID','LastInbound','AvgInbound','MaxInbound','LastOutbound','AvgOutbound','MaxOutbound']
    return dt_3
    #dt_3.to_excel(sheetbook, filename)
    #dt_3.to_excel("xmltest.xlsx", filename)

    #test
    '''
    srcfilenm = 'E:/tsun/ct-summary/test' + '/cacti_report_-2018-AUG.xml'
    df = pyxml_ms2003(srcfilenm, df_mapper)
    print(df)
   '''


def pyxml_ms2003(srcfile):
    # 从xml中提取有效数据，放入数组xmlist缓存
    #xmlist = []
    xml = objectify.parse(srcfile)
    root = xml.getroot()
    '''
    for element in root.iter():
        xmlist.append(element.text)
        #print("%s >>> %s" % (element.tag, element.text))
    '''
    xmlist = [ element.text for element in root.iter() ]
    # 数组定位重新设计
    bix = xmlist.index('00:00:00') - 10
    #bix

    xmlist.reverse()
    #xmlist

    #len(xmlist)
    eix = xmlist.index('00:00:00') - 17
    #eix

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

    # 运算有效数据矩阵存放
    res = []
    valid_cols = 7

    for row_c in range(icnt):
        for col_c in range(valid_cols):
            if ('-' not in idt.values[row_c][col_c]):
                #print (calMega(idt_2.values[row_c][col_c]))
                res.append(str(calMega(idt.values[row_c][col_c])) + str('M'))
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
    return dt_3
    #dt_3.to_excel(sheetbook, filename)
    #dt_3.to_excel("xmltest.xlsx", filename)


def pyxlsx_to_html(srcxlsx, srcsheet, tarsheethtml):
    df1 = pd.read_excel(srcxlsx, sheetname=srcsheet)
    strs = ['<HTML>']
    strs.append('<HEAD><TITLE>tarsheethtml</TITLE></HEAD>')
    strs.append('<BODY>')
    strs.append(df1.to_html())
    strs.append('</BODY></HTML>')
    html = "".join(strs)
    file = open(tarsheethtml, 'w')
    file.write(html)
    file.close()



def pyxml3_hyper_ms2003(srcfile):
    xml = objectify.parse(srcfile)
    root = xml.getroot()
    xmlist = [ e.text for e in root.iter() ]
    bix = xmlist.index('00:00:00') - 10
    xmlist.reverse()
    eix = xmlist.index('00:00:00') - 17
    xmlist.reverse()
    fcnt = 27
    cols = ['f1','f2','f3','f4','f5','f6','f7','f8','f9','f10','f11','f12','f13','f14','f15','f16','f17','f18','f19','f20','f21','f22','f23','f24','f25','f26','f27']
    df = pd.DataFrame()
    bbix = int(bix)
    eeix = len(xmlist) - eix - 1
    a = np.array(xmlist[bbix:eeix])
    icnt = int(len(xmlist[bbix:eeix])/fcnt)
    b = a.reshape(icnt, fcnt)
    sess = [pd.Series(b[i]) for i in range(len(b))]
    df = pd.concat(sess, axis = 1)
    df.index = cols
    dt = df.T
    idt = pd.DataFrame(dt, columns=['f3', 'f17', 'f19', 'f21', 'f23', 'f25', 'f27'])
    #idt
    #idt.shape
    valid_cols = 7
    res = [ 
        str(calMega(idt.values[row_c][col_c])) if ('-' not in idt.values[row_c][col_c]) else str(idt.values[row_c][col_c]) for row_c in range(icnt) for col_c in range(valid_cols) 
    ]
    #res
    o = np.array(res)
    out = o.reshape(icnt, valid_cols)
    sess_3 = [
       pd.Series(out[j]) for j in range(len(out))
    ]
    df_3 = pd.concat(sess_3, axis=1)
    #df_3
    dt_3 = df_3.T
    dt_3.columns.values
    dt_3.columns = ['UserID','LastInbound','AvgInbound','MaxInbound','LastOutbound','AvgOutbound','MaxOutbound']
    return (dt_3)


if __name__=='__main__':

    # Script starts from here
    print (sys.argv)
    if len(sys.argv) < 2:
        print ('No action specified.')
        sys.exit()
    if sys.argv[1].startswith('--'):
        option = sys.argv[1][2:]
        # fetch sys.argv[1] but without the first two characters
        if option == 'version':  #当命令行参数为-- version，显示版本号
            print ('Version 1.0 (Powered by TSunx)')
        elif option == 'help':  #当命令行参数为--help时，显示相关帮助内容
            print ('''
          design: Chen.Chaoqun
          
          Read MS 2003 xml file.
          This function is to execute the Mega unit output.     
          Any number of files can be specified .    
          Options include:    
          --all     : Feed all the .xml files in the TEST folder
          --file    : Point to a .xml file in the TEST folder
          --map     : Index the SN from the Mapper
          --html    : Make the multifile the HTML
          --hyper   : Speed up to generate the result
          --version : Prints the version number
          --help    : Display this help''')
        elif option == "file":
            print ('option: file')
            print ('xmlfile.xlsx to save the result.')
            writer = pd.ExcelWriter('xmlfile.xlsx')
            #print sys.argv[2]
            if sys.argv[2] in srcfiles:
                f = sys.argv[2]
                print (f)
                print ('file was read: %s' % f)
                srcfilenm = './test/' + f
                df = pyxml_ms2003(srcfilenm)
                df.to_excel(writer, sheet_name = f)
            else:
                print ('The %s NOT found in TEST folder!' % sys.argv[2])
        elif option == "all":
            print ('option: all')
            print ('xmlmultifile.xlsx to save the result.')
            writer = pd.ExcelWriter('xmlmultifile.xlsx')
            for f in srcfiles:
                print ('file was read: %s' % f)
                srcfilenm = './test/' + f
                df = pyxml_ms2003(srcfilenm)
                df.to_excel(writer, sheet_name = f)
            writer.save()
        elif option == "map":
            print ('option: map')
            print ('To index the mapper file!')
            print (mapperfile)
            writer = pd.ExcelWriter('xmlmapmultifile.xlsx')
            for f in srcfiles:
                print('file was read: %s' % f)
                srcfilenm = './test/' + f
                df = pyxmlm_ms2003(srcfilenm, df_mapper)
                df.to_excel(writer, sheet_name=f)
            writer.save()
        elif option == "html":
            print('option: html')
            print('Make the multifile the HTML!')
            for f in srcfiles:
                # pyxlsx_to_html(srcxlsx, srcsheet, tarsheethtml)
                tarsheethtmlnm = './' + f + '.html'
                pyxlsx_to_html('xmlmultifile.xlsx', f, tarsheethtmlnm)
                print ('%s created.' % tarsheethtmlnm)
        elif option == "hyper":
            print ('option: hyper')
            print ('xmlmultifile.xlsx to save the result.')
            writer = pd.ExcelWriter('xmlmultifile.xlsx')
            for f in srcfiles:
                print ('file was read: %s' % f)
                srcfilenm = './test/' + f
                df = pyxml3_hyper_ms2003(srcfilenm)
                df.to_excel(writer, sheet_name = f)
            writer.save()
        else:
            print ('Unknown option.')
            sys.exit()

