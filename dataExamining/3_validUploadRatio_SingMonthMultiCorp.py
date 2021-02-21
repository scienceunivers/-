import os, sys
import pandas as pd, numpy as np
import datetime as dt
import time
import random
# import xlrd, openpyxl

'''
workbook_successedUploadRatio = xlrd.open_workbook(r'D:\...\gxxp\20190906\4-1 ajgrkl(2019-09-16)_复制粘贴.xlsx')
sheetNames = workbook_successedUploadRatio.sheet_names()
print(sheetNames)
objectiveSheet = sheetNames[0]

# 现在的生产系统计算逻辑是不现实的。

# data['当月入库率'] = (data['入库记录数'] - data['错误记录显示数'])/data['入库记录数']
data['入库率'] = data['入库率'].str.strip('%').astype('float')
# https://stackoverflow.com/questions/50686004/change-column-with-string-of-percent-to-float-pandas-dataframe
successedUploadRatio = data.copy()
print(successedUploadRatio.shape)
'''

updateDetail = pd.read_csv(r'D:\...\gxpt\20200221\202002281500mxbsqkck.csv',
                           sep='\s*,', header=2, 
                           
                           index_col=None, usecols=None,
                           dtype=None, engine='python',
                           skipinitialspace=False, skiprows=None, skipfooter=0, nrows=None,
                           na_values=None, keep_default_na=True, na_filter=True,skip_blank_lines=True, 
                           parse_dates=False, infer_datetime_format=False, keep_date_col=False, date_parser=None, dayfirst=False, 
                           cache_dates=True, 
                           iterator=False, chunksize=None, compression='infer', thousands=None, decimal=b'.', 
                           lineterminator=None, quotechar='"', quoting=0, doublequote=True, escapechar=None, 
                           comment=None, encoding=None, dialect=None, 
                           error_bad_lines=True, warn_bad_lines=True, delim_whitespace=False, 
                           low_memory=True, memory_map=False, float_precision=None)
# names=['序号','文件名称','文件类型','报送机构','报送时间','入库时间','处理状态','入库记录数','出错记录数'], 

# xlrd读csv:
# XLRDError: Unsupported format, or corrupt file: Expected BOF record; found b'\xb2\xe9\xd1\xaf\xcc\xf5\xbc\xfe',
# 对这个错误的分析见External-excel processing-xlrd, xlwt.docx。


# 明细报送，得先清洗数据，有些记录没有入库时间，导致后面几列错位，excel的话从最后一列出错记录开始修正
# 用read_csv，sep=None提示Error: Could not determine delimiter。

# pd.read_csv(r'D:\...\20191229\20191229mxbsqkck.csv', sep='[\t,]')
# ParserError: Expected 1 fields in line 3, saw 9. Error could possibly be due to quotes being ignored
# when a multi-char delimiter is used.
# pd.read_csv(r'D:\...\20191229\20191229mxbsqkck.csv', sep=……, 
#             delimiter=None, header=2, names=None, index_col=0, usecols=None, dtype=None, engine='python')

# ’(\s)*,’，0或多次whitespace+逗号，但匹配的结果总是会在9个字段两两之间多出8个字段，
# 字段名是none，字段的内容都是清一色\t。
# 也不是regex没有转义的原因，与是否greedy无关。
# 肯定是括号的原因。

# http://www.regexplanet.com/advanced/python/index.html给’(\s)*,’的反馈是()出错，
# http://regex.larsolavtorvik.com/给’(\s)*,’的反馈是匹配出了2个数组，array1是8个逗号，array2是8个none，
# 引擎用的是PHP的preg_match_all


grouped = updateDetail.groupby('报送机构').agg({'入库记录数':sum,'出错记录数':sum})
# group完了接着针对2列汇总。
successedUploadRatio = grouped.assign(入库率=grouped['入库记录数']/(grouped['出错记录数']+grouped['入库记录数']))
successedUploadRatio_replaced = successedUploadRatio.replace(to_replace=np.nan,value='#DIV/0!')


with pd.ExcelWriter(r'D:...\gxpt\20200907\{:%Y%m%d%H%M}successedUploadRatio.xlsx'.format(dt.datetime.today()),
                       date_format='YYYY/MM/DD') as writer:
    successedUploadRatio_replaced.to_excel(writer,
                                         sheet_name='successedUploadRatio_replaced',
                                         
                     )

# 后续考虑用pickle，或%store，实现share date between notebooks。
