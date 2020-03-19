import os
import pandas as pd, numpy as np
import sys
import datetime as dt
import time
import random
import xlrd


workbook_dataPerMonth = xlrd.open_workbook(r'D:\...\gxpt\20200221\1-1gyywl(2020-03-03).xls')
sheetNames = workbook_dataPerMonth.sheet_names()
print(sheetNames)
objectiveSheet = sheetNames[0]
# workbook_dataPerMonthdata = pd.DataFrame(workbook_dataPerMonth) 不能直接转化成df
nRows, nCols = workbook_dataPerMonth.sheet_by_name(objectiveSheet).nrows, workbook_dataPerMonth.sheet_by_name(objectiveSheet).ncols
colNames = []
workbook_dataPerMonthDict = {}
for j in range(0,nRows):
    if any(workbook_dataPerMonth.sheet_by_name(objectiveSheet).row_values(j)):
        print(j)
        break
for i in range(0, nCols):
    colNames.append(workbook_dataPerMonth.sheet_by_name(objectiveSheet).col_values(i)[j])
    workbook_dataPerMonthDict[workbook_dataPerMonth.sheet_by_name(objectiveSheet).col_values(i)[j]] = \
        workbook_dataPerMonth.sheet_by_name(objectiveSheet).col_values(i)[j+1:]
data = pd.DataFrame(workbook_dataPerMonthDict, columns=colNames)



# 如果用openpyxl，可以考虑写一个将xls数据复制到xlsx的脚本。

# 直接xlrd，会把前面的14行空白行读取过来，原因在于表名、单位（比如各月业务量、单位）放在文本框里

print(data.shape)

consecutiveIndicator=[]
for index in data.index:
    row = data.loc[index,:]
    for month in range(1,data.loc[:,'2016年10':].shape[1]+1):
        if row[-month]!=0:
            continue
        else:
            break
    if month<data.loc[:,'2016年10':].shape[1]:
        consecutiveIndicator.append(month-1)
    else:
        consecutiveIndicator.append(month)
# consecutiveIndicator = (data.iloc[:,-3]!=0)&(data.iloc[:,-2]!=0)&(data.iloc[:,-1]!=0)
# # 用bitwise &，关于and和&，见reference和
# # 22646463/and-boolean-vs-bitwise-why-difference-in-behavior-with-lists-vs-nump
data.insert(5,'lxbsyfs',consecutiveIndicator)
dataPerMonth = data.copy()
print(dataPerMonth.shape)
