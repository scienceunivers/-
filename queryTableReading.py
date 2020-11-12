import os, sys
import pandas as pd, numpy as np
import datetime as dt
import time
import random
import xlrd
# openpyxl xlwt xlwriter xlrd等的比较，见
# https://zhuanlan.zhihu.com/p/23998083
# https://www.pythonexcel.com/
# openpyxl, 功能全面，可读可写，对格式也有一定的操作，xlwriter对格式操作支持较好，pandas也可以写excel。

# 读取上一期查询总表
workbook_query = xlrd.open_workbook(
    r'D:\...\5-查询接入总表202004031650-正式进入生产环境名单404+43.xlsx')
sheetNames_query = workbook_query.sheet_names()
print(sheetNames_query)
objectiveSheet = '接入机构0331'
# 之后 objectiveSheet 可以作为参数传进来。
# workbook_querydata = pd.DataFrame(workbook_query) 不能直接转化成df
nRows, nCols = workbook_query.sheet_by_name(objectiveSheet).nrows, workbook_query.sheet_by_name(objectiveSheet).ncols
colNames = []
workbook_queryDict = {}
for i in range(0, nCols):
    colNames.append(workbook_query.sheet_by_name(objectiveSheet).col_values(i)[0])
    workbook_queryDict[workbook_query.sheet_by_name(objectiveSheet).col_values(i)[0]] = \
        workbook_query.sheet_by_name(objectiveSheet).col_values(i)[1:]
data = pd.DataFrame(workbook_queryDict, columns=colNames)
# 将excel中的错误值保持原样。
for i in range(1, nRows):
    for j in range(0, nCols):
        if workbook_query.sheet_by_name(objectiveSheet).cell_type(i, j) == 5:
            if workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == 7:
                data.iloc[i-1, j] = '#DIV/0!'
            elif workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) != 42:
                # 保证后面只出现NA类型的错误。
                raise Exception('Redundant kinds of errors in workbook_queryDict at [{},{}].'.format(i, j))
                # 20190212以后超过两层括号考虑分行写，否则少打一个括号难以排查。
            else:
                data.iloc[i-1, j] = '#N/A'
        # 查询接入总表中有DIV，也有NA，以后如果还有其他的错误，即cell type =5，继续套用elif的框架，加进去即可。
        if workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == '':
            #data.iloc[i - 1, j] = None
            #data.iloc[i - 1, j] = np.nan
            # https://stackoverflow.com/questions/17534106/what-is-the-difference-between-nan-and-none
            # nan与None的区别。
            data.iloc[i - 1, j] = ''
            # 如果不是数值，这个语句也不用加，但保险起见，还是加上。
        if j<(nCols-3) and workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == 0:
            data.iloc[i - 1, j] = ''
            # 还是以保真为原则，先不考虑速度，所以不用nan。
            # 因为vlookup把空弄过来自动填为0，但xlrd读取的时候还是会读成0。
            
# xlrd读过来的日期是数字，excel日期和数字的转换，见xlrd的摘录。

print(data.shape,
     data.loc[data['接入状态']=='',:].shape)
queryTable_all = data.copy()

queryTable_all.head(100)

# 20191202test.assign(开通查询时间=test.apply(xlrd.xldate.xldate_as_datetime, datemode=0))
# 用series版的apply就可以。df版的apply就报错，在if xldate < 60:这一步，错误是
# ValueError: ('The truth value of a Series is ambiguous. Use a.empty, a.bool(), a.item(), a.any() or a.all().', 
# 'occurred at index 开通查询时间')
# 我看了一下，虽然pandas api reference中对df.apply的解释是Apply a function along an axis of the DataFrame. 但没说清是对
# 整个axis做运算还是按顺序对axis中的元素做运算，举例的时候举的也是不区分整列还是element wise的，比如sqr、+1这种运算。
# https://stackoverflow.com/questions/34962104/pandas-how-can-i-use-the-apply-function-for-a-single-column
# 其中一个回答下的评论：map() is for Series (i.e. single columns) and operates on one cell at a time, 
# while apply() is for DataFrame, and operates on a whole row at a time. – jpcgt，这个是关键。所以一次性把
# 一列传到xldate_as_datetime中xldate < 60是无法判断的。
# map也得是series才可以，series.apply倒是一整列一起运算、element wise都行。
# queryTable_all.loc[(queryTable_all['开通查询时间']!='#N/A')&(queryTable_all['开通查询时间']!=''),'开通查询时间']=
# test.apply(xlrd.xldate.xldate_as_datetime, datemode=0),直接赋值后，出现了非常怪异的值，变成了一长串比如1547424000000000000
# 这也不是julian dates。
# 所以，提取series，用apply，可以保证datemode被传递，提取series用map，datemode没法体现。
# 提取df，用applymap，datemode保证不了，提取df用assign，会返回新的object。
# 而且由于都是对筛选出来的部分做操作，反应不到原始df queryTable_all上。当然由于筛选出来的series或df，index都没变，
# 所以可以考虑循环赋值。df.where()也没用，一样得先筛选出开通查询时间这一列。
# 像https://stackoverflow.com/questions/38454403/convert-excel-style-date-with-pandas#的某个答案一样，
# 用pd.to_datetime+xldate_as_datetime也得先把数字筛出来，不筛xldate_as_datetime报错。最后还得用循环去赋值。
# seriesBeforeTransform = queryTable_all.loc[(queryTable_all['开通查询时间']!='#N/A')&(queryTable_all['开通查询时间']!=''),'开通查询时间']
# dfAfterTransform = queryTable_all.loc[(queryTable_all['开通查询时间']!='#N/A')&(queryTable_all['开通查询时间']!=''),['开通查询时间']].\
# assign(开通查询时间=beforeTransform.apply(xlrd.xldate.xldate_as_datetime, datemode=0))
# for index in dfAfterTransform.index:
#     queryTable_all.loc[index,'开通查询时间']=dfAfterTransform.loc[index,'开通查询时间']
    
# 20191209seriesBeforeTransform = queryTable_all.loc[(queryTable_all['更新日期']!='#N/A')&(queryTable_all['更新日期']!=''),'更新日期']
# dfAfterTransform=queryTable_all.loc[(queryTable_all['更新日期']!='#N/A')&(queryTable_all['更新日期']!=''),['更新日期']].\
# assign(更新日期=beforeTransform.apply(xlrd.xldate.xldate_as_datetime, datemode=0))
# for index in dfAfterTransform.index:
#     print(dfAfterTransform.loc[index,'更新日期'])
# 本来以为这样就可以了，开通查询时间已经通过验收，但更新日期却出现了NaT，excel中的正常的日期，跟其他日期一样都是公式vlookup
# 过来的，不知道为什么用xldate_as_datetime转换的时候出现了NaT，筛选非空非#NA记录的时候，筛出来106个数字，没有异常，
# xldate_as_datetime转换完就有NaT了。这个问题暂时无解。
# 还是得用https://stackoverflow.com/questions/38454403/convert-excel-style-date-with-pandas#的第一个答案。

# seriesBeforeTransform = queryTable_all.loc[(queryTable_all['开通时间']!='#N/A')&(queryTable_all['开通时间']!=''),'开通时间']
# dfAfterTransform = queryTable_all.loc[(queryTable_all['开通时间']!='#N/A')&(queryTable_all['开通时间']!=''),['开通时间']].\
# assign(开通时间=beforeTransform.apply(xlrd.xldate.xldate_as_datetime, datemode=0))
# for index in dfAfterTransform.index:
#     queryTable_all.loc[index,'开通时间']=dfAfterTransform.loc[index,'开通时间']
df1=queryTable_all.loc[(queryTable_all['开通时间']!='#N/A')&(queryTable_all['开通时间']!=''),['开通时间']]
df1['开通时间'] = pd.TimedeltaIndex(df1['开通时间'], unit='d') + dt.datetime(1899,12,30)
for index in df1.index:
    queryTable_all.loc[index,'开通时间']=df1.loc[index,'开通时间'].date()
    # datetime.date()不能用于series只能写循环了。
print(queryTable_all.loc[(queryTable_all['开通时间']!='#N/A')&(queryTable_all['开通时间']!=''),['开通时间']].shape)

# 20191214目前的问题就在于df1['开通查询时间'] = pd.TimedeltaIndex(df1['开通查询时间'], unit='d') + dt.datetime(1899,12,30)
# 计算出来的是yyyy-mm-dd格式，但赋值给queryTable_all后就成了yyyy-mm-dd hh:mm:ss格式。

# 执行df1['开通查询时间'].dtypes
# 显示<M8[ns]，查了一下，https://stackoverflow.com/questions/29206612/difference-between-data-type-datetime64ns-and-m8ns
# datetime64[ns] is a general dtype, while <M8[ns] is a specific dtype. General dtypes map to specific dtypes, 
# but may be different from one installation of NumPy to the next.
# 然后说明了由于计算机体系是big endian或little endian，会导致映射为<M8[ns]还是>M8[ns]两种不同类型，这个问题。
# 对于什么是big endian或little endian，https://blog.csdn.net/waitingbb123/article/details/80504093讲得比较清楚了。

# 20191215执行pd.TimedeltaIndex(df1['开通查询时间'], unit='d') + dt.datetime(1899,12,30)
# 提示FutureWarning: Passing datetime64-dtype data to TimedeltaIndex is deprecated, will raise a TypeError in a future version
# 原因在于缓存中的df1没有删除，还是yyyy-mm-dd的状态，类型是datetime64，所以说Passing datetime64-dtype data to 
# TimedeltaIndex is deprecated，所以输出结果也都是1949年的。

# 恢复被删除的cell：https://stackoverflow.com/questions/48814309/how-to-recover-deleted-jupyter-notebook-cell

# 20191220如果我用df_test['开通查询时间'] = pd.TimedeltaIndex(df_test['开通查询时间'], unit='d') + dt.date(1899,12,30)，提示
# TypeError: unsupported operand type(s) for +: 'TimedeltaIndex' and 'datetime.date'。
# 这个错误意思是这两种类型不能相加。
# df1['开通查询时间'] = pd.TimedeltaIndex(df1['开通查询时间'], unit='d') + dt.datetime(1899,12,30)执行完
# 后，显示的是DatetimeIndex(['2019-01-14'……，年月日的datetimeindex，但数据类型是'datetime64[ns]'，元素依然是datetime
# 显示的时候没有时间部分，但赋值后就带有了时间部分，他妈的不知道为什么。
# 直接用instance method datetime.date()解决！也不必datetime.date().strftime多此一举。

df2=queryTable_all.loc[(queryTable_all['更新日期']!='#N/A')&(queryTable_all['更新日期']!=''),['更新日期']]
df2['更新日期'] = pd.TimedeltaIndex(df2['更新日期'], unit='d') + dt.datetime(1899,12,30)
for index in df2.index:
    queryTable_all.loc[index,'更新日期']=df2.loc[index,'更新日期'].date()
print(queryTable_all.loc[(queryTable_all['更新日期']!='#N/A')&(queryTable_all['更新日期']!=''),['更新日期']].shape)
# 日志用logging还不如print。
