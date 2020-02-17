import os, sys
import pandas as pd, numpy as np
import datetime as dt
import time
import random
import xlrd

workbook_query = xlrd.open_workbook(r'D:\hlwjrxh\mqdgz\gxpt\cyzb\5-cxjrzb202001042031.xlsx')
sheetNames_query = workbook_query.sheet_names()
print(sheetNames_query)
objectiveSheet = '接入机构'
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
            else:
                data.iloc[i-1, j] = '#N/A'
        # 查询接入总表中有DIV，也有NA，以后如果还有其他的错误，即cell type =5，继续套用elif的框架即可。
        if workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == '':
            #data.iloc[i - 1, j] = None
            #data.iloc[i - 1, j] = np.nan
            data.iloc[i - 1, j] = ''
            # 如果不是数值，这个语句也不用加，但保险起见，还是加上。
        if j<(nCols-3) and workbook_query.sheet_by_name(objectiveSheet).cell_value(i, j) == 0:
            data.iloc[i - 1, j] = ''
            # 还是以保真为原则，先不考虑速度，所以不用nan。
            # 因为vlookup把空弄过来自动填为0，但xlrd读取的时候还是会读成0。
            
print(data.shape)
queryTable_all = data.copy()
