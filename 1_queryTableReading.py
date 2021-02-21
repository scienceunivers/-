import os, sys
import pandas as pd, numpy as np
import datetime as dt
import time
import random
import xlrd
# openpyxl xlwt xlwriter xlrd等的比较，见
# https://zhuanlan.zhihu.com/p/23998083
# https://www.pythonexcel.com/

# 读取上一期查询总表
workbook_query = xlrd.open_workbook(
    r'D:\...\5-查询接入总表202004031650-正式进入生产环境名单404+43.xlsx')
sheetNames_query = workbook_query.sheet_names()
print(sheetNames_query)
objectiveSheet = '接入机构0331'
# 之后 objectiveSheet 可以作为函数参数传进来。
nRows, nCols = workbook_query.sheet_by_name(objectiveSheet).nrows, workbook_query.sheet_by_name(objectiveSheet).ncols
colNames = []
workbook_queryDict = {}
for i in range(0, nCols):
    colNames.append(workbook_query.sheet_by_name(objectiveSheet).col_values(i)[0])
    workbook_queryDict[workbook_query.sheet_by_name(objectiveSheet).col_values(i)[0]] = \
        workbook_query.sheet_by_name(objectiveSheet).col_values(i)[1:]
data = pd.DataFrame(workbook_queryDict, columns=colNames)
# 由于xlrd的局限，需加代码将错误值保持原样。
for i in range(1, nRows):
    for j in range(0, nCols):
        if workbook_query.sheet_by_name(objectiveSheet).cell_type(i, j) == 5:
            if workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == 7:
                data.iloc[i-1, j] = '#DIV/0!'
            elif workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) != 42:
                # 保证后面只出现NA类型的错误。
                raise Exception('Redundant kinds of errors in workbook_queryDict at [{},{}].'.format(i, j))
                
            else:
                data.iloc[i-1, j] = '#N/A'
        # 查询接入总表中有DIV，也有NA，以后如果还有其他的错误，即cell type =5，继续套用elif的框架，加进去即可。
        if workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == '':

            data.iloc[i - 1, j] = ''
            # 如果不是数值，这个语句也不用加，但保险起见，还是加上。
        if j<(nCols-3) and workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == 0:
            data.iloc[i - 1, j] = ''
            # 以保真为原则，先不考虑速度，所以不用nan。
            # 因为vlookup把空弄过来自动填为0，但xlrd读取的时候还是会读成0。
            
# xlrd读过来的日期是数字，需要增加excel日期和数字的转换。

print(data.shape,
     data.loc[data['接入状态']=='',:].shape)
queryTable_all = data.copy()

queryTable_all.head(100)


# 注意几个陷阱：

# ValueError: ('The truth value of a Series is ambiguous. Use a.empty, a.bool(), a.item(), a.any() or a.all().', 
# 'occurred at index 开通查询时间')

# queryTable_all.loc[(queryTable_all['开通查询时间']!='#N/A')&(queryTable_all['开通查询时间']!=''),'开通查询时间']=
# test.apply(xlrd.xldate.xldate_as_datetime, datemode=0),直接赋值后，出现了非常怪异的值，比如1547424000000000000
# 这不是数字格式日期，也不是julian dates。

# 提取series，用apply，可以保证datemode被传递，提取series用map，datemode没法体现。
# 提取df，用applymap，datemode保证不了，提取df用assign，会返回新的object。
# 由于都是对筛选出来的部分做操作，反应不到原始df queryTable_all上。当然由于筛选出来的index都没变，
# 所以可以考虑循环赋值。

df1=queryTable_all.loc[(queryTable_all['开通时间']!='#N/A')&(queryTable_all['开通时间']!=''),['开通时间']]
df1['开通时间'] = pd.TimedeltaIndex(df1['开通时间'], unit='d') + dt.datetime(1899,12,30)
for index in df1.index:
    queryTable_all.loc[index,'开通时间']=df1.loc[index,'开通时间'].date()
    # datetime.date()不能用于series只能写循环了。
print(queryTable_all.loc[(queryTable_all['开通时间']!='#N/A')&(queryTable_all['开通时间']!=''),['开通时间']].shape)

# 20191214目前的问题就在于df1['开通查询时间'] = pd.TimedeltaIndex(df1['开通查询时间'], unit='d') + dt.datetime(1899,12,30)
# 计算出来的是yyyy-mm-dd格式，但赋值给queryTable_all后就成了yyyy-mm-dd hh:mm:ss格式。
# df1['开通查询时间'].dtypes
# 显示<M8[ns]，计算机体系是big endian或little endian，会导致映射为<M8[ns]还是>M8[ns]两种不同类型，这个问题。

# 20191215执行pd.TimedeltaIndex(df1['开通查询时间'], unit='d') + dt.datetime(1899,12,30)
# 提示FutureWarning: Passing datetime64-dtype data to TimedeltaIndex is deprecated, will raise a TypeError in a future version
# 原因在于缓存中的df1没有删除，还是yyyy-mm-dd的状态，类型是datetime64，所以说Passing datetime64-dtype data to 
# TimedeltaIndex is deprecated，所以输出结果也都是1949年的。

df2=queryTable_all.loc[(queryTable_all['更新日期']!='#N/A')&(queryTable_all['更新日期']!=''),['更新日期']]
df2['更新日期'] = pd.TimedeltaIndex(df2['更新日期'], unit='d') + dt.datetime(1899,12,30)
for index in df2.index:
    queryTable_all.loc[index,'更新日期']=df2.loc[index,'更新日期'].date()
print(queryTable_all.loc[(queryTable_all['更新日期']!='#N/A')&(queryTable_all['更新日期']!=''),['更新日期']].shape)
# 日志用logging不如print。
