import os
import pandas as pd, numpy as np
import sys
import datetime as dt
import time
import random
import xlrd

workbook_updateRatio = xlrd.open_workbook(r'D:\...\gxpt\20200221\3-1gygxlxx(2020-03-03).xls',
                                         formatting_info=True)
sheetNames = workbook_updateRatio.sheet_names()
print(sheetNames)
objectiveSheet = sheetNames[0]

nRows, nCols = workbook_updateRatio.sheet_by_name(objectiveSheet).nrows, workbook_updateRatio.sheet_by_name(objectiveSheet).ncols
colNames = []
colNames_Row1,colNames_Row2 = [],[]
workbook_updateRatioDict = {}
for j in range(0,nRows):
    if any(workbook_updateRatio.sheet_by_name(objectiveSheet).row_values(j)):
        print(j)
        break
for i in range(0, nCols):
    colNames.append(workbook_updateRatio.sheet_by_name(objectiveSheet).col_values(i)[j])
    workbook_updateRatioDict[workbook_updateRatio.sheet_by_name(objectiveSheet).col_values(i)[j]] = \
        workbook_updateRatio.sheet_by_name(objectiveSheet).col_values(i)[j+1:]
        # 为了直接将excel raw data读入，生成装有raw data的df，不管被merge的表头。
        # 0217但这样不行，因为第一行的合并单元格中，有''作为列名，会重复，重复的列在df中会保留，但重复的键值在dict中会被覆盖，所以
        # 导致dict中键值数小于colNames长度，再生成df的时候会有一些列没有值，默认从上列复制过来，所以出现很多列都是有效+G的情况。



    # 原则应该是先把互为笛卡尔积的level找出来，月份与有效类型是，这两者与序号等几列也可以互为笛卡尔积。
    colNames_Row1.append(workbook_updateRatio.sheet_by_name(objectiveSheet).col_values(i)[j])
    colNames_Row2.append(workbook_updateRatio.sheet_by_name(objectiveSheet).col_values(i)[j+1])
    # 以后可以考虑把j+1一般化为j+k，容纳表头为3行及以上单元格合并而成的情况。
    
data = pd.DataFrame(workbook_updateRatioDict, columns=colNames)
# 0217这样用dict生成dataframe不行，因为第一行的合并单元格中，有''作为列名，会重复，重复的列在df中会保留，但重复的键值在dict中会被覆盖，所以
# 导致dict中键值数小于colNames长度，再生成df的时候会有一些列没有值，默认从上列复制过来，所以出现很多列都是有效+G的情况。
# 还是得用read_excel去把业务数据这一部分作为一个multiIndex单独读取出来，再concat基本信息。

# 将前面几列抽出来。
colNames_Row1,colNames_Row2 = pd.Series(colNames_Row1),pd.Series(colNames_Row2)
standAloneLabelIndex = []
for k in range(len(colNames_Row1)):
    if colNames_Row1[k] != '' and colNames_Row2[k] == '':
        standAloneLabelIndex.append(k)
# standAloneLabel = colNames_Row1[standAloneLabelIndex] 
standAloneLabel = [colNames_Row1[k] for k in range(len(colNames_Row1)) if colNames_Row1[k] != '' and colNames_Row2[k] == '']
print(standAloneLabel)


# 0208处理level1，生成月份维度，list
level1 = colNames_Row1.copy()
for index in standAloneLabelIndex:
    level1.pop(index)
# level1Index = []
# for m in range(len(level1)):
#     if level1[m+index+1] != '':
#         # level1是series，被pop之后，被pop的元素的index也会被pop掉，所以要+index+1
#         level1Index.append(m+index+1)
# cleanLevel1 = level1[level1Index]
cleanLevel1 = [level1[m+index+1] for m in range(len(level1)) if level1[m+index+1] != '']
# list comprehension也可以，有个好处是自动reindex
# 注意这个list comprehension依赖变量index的最大值。
print(cleanLevel1)

# 处理level2，生成有效与否维度，list。
cleanLevel2 = set(colNames_Row2)
cleanLevel2.remove('')
# remove return None，与list的pop一样，但set的pop会返回被pop的东西。
cleanLevel2 = list(cleanLevel2)
print(cleanLevel2)

# 构造multi index df。
basicInfo = ['basicInfo']
updateTableIndex = pd.MultiIndex.from_product(iterables=[cleanLevel1,cleanLevel2],
                                   names=['month','DataType'])

# 怎么将excel中间的一堆数提取出来的，container选哪个？

updateTableFromReadExcel = pd.read_excel(r'D:\...\gxpt\20200221\3-1gygxlxx(2020-03-03).xls',
                sheet_name=0,header=[j,j+1], names=None, index_col=None, 
                usecols=None, squeeze=False, dtype=None, engine=None, 
                converters=None, true_values=None, false_values=None, skiprows=None, nrows=None, 
                na_values=None, keep_default_na=True, verbose=False, 
                parse_dates=False, date_parser=None, thousands=None, comment=None, 
                skipfooter=0, convert_float=True, mangle_dupe_cols=True)
# 如果在use_col中限定了具体的哪些列，则提示cannot specify usecols when specifying a multi-index header。
# https://github.com/pandas-dev/pandas/issues/25449
# timeDataType = updateTableFromReadExcel.copy().loc[:,cleanLevel1[0]:cleanLevel1[-1]]
# UnsortedIndexError: 'Key length (1) was greater than MultiIndex lexsort depth (0)'

# https://github.com/pandas-dev/pandas/issues/19771

# updateTableFromReadExcel.sort_index(axis=1,level=None,ascending=False,sort_remaining=True,inplace=True)
# 在这里倒序排序，鸟用没有，后面还得正序排序。


updateTableFromReadExcel.fillna(value={('管理员账号','Unnamed: 3_level_1'):''},method=None,axis=None,inplace=True)





# 计算df_effectiveUnpaid。
# updateTableFromReadExcel.assign((slice(None),'df_effectiveUnpaid')=(slice(None),'df_effective')-(slice(None),'df_effectivePaid')
#                                       -(slice(None),'Z')-(slice(None),'G')-(slice(None),'D'))
# SyntaxError: keyword can't be an expression，(slice(None)
# 57950935/pandas-df-assign-does-not-work-with-variable-names，提示可以用unpack：
# updateTableFromReadExcel.assign(**{(slice(None),'df_effectiveUnpaid'):updateTableFromReadExcel.loc[:,(slice(None),'df_effective')]
#                                           -updateTableFromReadExcel.loc[:,(slice(None),'df_effectivePaid')]})
# TypeError: unhashable type: 'slice'
# for month in cleanLevel1:
#     updateTableFromReadExcel=updateTableFromReadExcel.assign((month,df_effectiveUnpaid)=
#                                                                               updateTableFromReadExcel.loc[:,(month,'df_effective')]
#                                                                               -updateTableFromReadExcel.loc[:,
#                                                                                                                    (month,'df_effectivePaid')
#                                                                                                                   ]
#                                                                              )
# SyntaxError: keyword can't be an expression。
# 还是暴力计算吧。

df_effective = updateTableFromReadExcel.loc[:,(slice(None),'yxbs')].rename({'yxbs':'unified'},axis=1,level=1)
df_effectivePaid = updateTableFromReadExcel.loc[:,(slice(None),'yxjqbs')].rename({'yxjqbs':'unified'},axis=1,level=1)
df_effectiveZ = updateTableFromReadExcel.loc[:,(slice(None),'Z')].rename({'Z':'unified'},axis=1,level=1)
df_effectiveG = updateTableFromReadExcel.loc[:,(slice(None),'G')].rename({'G':'unified'},axis=1,level=1)
df_effectiveD = updateTableFromReadExcel.loc[:,(slice(None),'D')].rename({'D':'unified'},axis=1,level=1)
df_effectiveUnpaid = (df_effective-df_effectivePaid-df_effectiveZ-df_effectiveG-df_effectiveD).rename(
    {'unified':'yxwjq'},axis=1,level=1)
# 既然列名不同运算出来都是nan，那就改成一样的。最后join一下.
# 由于updateTableFromReadExcel是unsorted，会提示UnsortedIndexError。这时必须做一个正向排序。
updateTableWithUnpaid = updateTableFromReadExcel.join(df_effectiveUnpaid,on=None,how='inner').sort_index(
    axis=1,level=None,ascending=False,sort_remaining=True)


# 计算updateRatio。
# 必须做一个正向排序。
levelValues = list(set(updateTableWithUnpaid.sort_index(
    axis=1,level=0,ascending=True,sort_remaining=False).columns.get_level_values(0)
                       ) & set(cleanLevel1))
levelValues.sort()
addition_effective = updateTableWithUnpaid.sort_index(axis=1,level=0,ascending=True,sort_remaining=False).loc[
    :,(slice(levelValues[0],levelValues[-2]),'yxbs')].sum(axis=1,level=None)
addition_effectiveUnpaid = updateTableWithUnpaid.sort_index(axis=1,level=0,ascending=True,sort_remaining=False).loc[
    :,(slice(levelValues[0],levelValues[-2]),'yxwjq')].sum(axis=1,level=None)

updateRatioTable = pd.concat([updateTableWithUnpaid.loc[:,standAloneLabel].copy().droplevel(axis=1,level=1),addition_effective,
          addition_effectiveUnpaid],axis=1).rename({0:'yxbs',1:'yxwjq'},axis=1)
# 看来df、series混起来也可以。
# updateRatioTable.assign(UNupdateRatio='yxwjq'/'yxbs',updateRatio=1-UNupdateRatio)
updateRatioTable['UNupdateRatio']=updateRatioTable['yxwjq']/updateRatioTable['yxbs']
updateRatioTable['updateRatio']=1-updateRatioTable['UNupdateRatio']

# 这里也得replace一下，有DIV！。
updateRatioTable_replaced = updateRatioTable.replace(to_replace=np.nan,value='#DIV/0!')


with pd.ExcelWriter(r'D:\...\gxpt\20201109\{:%Y%m%d%H%M}updateRatioTable.xlsx'.format(dt.datetime.today()),
                       date_format='YYYY/MM/DD') as writer:
    updateTableWithUnpaid.to_excel(writer,
                                         sheet_name='updateTableWithUnpaid',
                                         na_rep='#N/A'
                     )
    updateRatioTable_replaced.to_excel(writer,
                                         sheet_name='updateRatioTable_replaced',
                                         na_rep='#N/A')
# updateTableWithUnpaid写进excel的时候会在表头下多一个空行，原因在于https://github.com/pandas-dev/pandas/issues/27772，
# 空行是为了以下情景：如果index是multiindex，而且多个level都被命名了，这个空行是给level的名字留的空间。这是个bug。

'''

# data_read_excel = pd.read_excel(r'D:\...\gxpt\20190906\201909291700 3-1gygxlxx(2019-09-29)_复制粘贴.xlsx',
#                                 sheet_name=0,header=None, names=None, index_col=None, 
#                                 usecols='HG:HM', squeeze=False, dtype=None, engine=None, 
#                                 converters=None, true_values=None, false_values=None, skiprows=None, nrows=None, 
#                                 na_values=None, keep_default_na=True, verbose=False, 
#                                 parse_dates=False, date_parser=None, thousands=None, comment=None, 
#                                 skipfooter=0, convert_float=True, mangle_dupe_cols=True)

# 先用excel column letter。

# header = data_read_excel.iloc[:2,:].ffill(axis=1).ffill(axis=0).iloc[1,:]
# print(header)

# body = pd.read_excel(r'D:\...\gxpt\20190906\201909291700 3-1gygxlxx(2019-09-29)_复制粘贴.xlsx',
#                    sheet_name=0,skiprows=[0,1], index_col=None, header=None,usecols='HG:HM')
# body.columns = header.values
# updateRatio = body.copy()




# https://github.com/pandas-dev/pandas/blob/master/pandas/io/common.py，源码显示
# na_values : scalar, str, list-like, or dict, default None
#     Additional strings to recognize as NA/NaN. If dict passed, specific
#     per-column NA values. By default the following values are interpreted
#     as NaN: '"""
#     + fill("', '".join(sorted(_NA_VALUES)), 70, subsequent_indent="    ")
#     + """'.
# _NA_VALUES在common.py中，如下：
# common NA values
# no longer excluding inf representations
# '1.#INF','-1.#INF', '1.#INF000000',
# _NA_VALUES = {
#     "-1.#IND","1.#QNAN","1.#IND","-1.#QNAN", "#N/A N/A","#N/A","N/A","n/a","NA","#NA","NULL","null","NaN","-NaN","nan","-nan","",
# }


'''
