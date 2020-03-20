import os
import pandas as pd, numpy as np
import sys
import datetime as dt
import time
import random
import xlrd

queryTable_all.columns

# 用全称vlookup。
result_1 = pd.merge(queryTable_all,
                  dataPerMonth.loc[:,['部门', '连续报送月']],
                  how='left', left_on='全称', right_on='部门',suffixes=('_0_L','_0_R'))
# result[result['连续报送月'].isna()].shape，这个验证或者日志用。
# 注意不能用np.nan==np.nan，见pandas 15.1.2 Values considered “missing”。
print(result_1.columns)
print(result_1[result_1['连续报送月'].isna()].shape)



# 用社会信用代码vlookup。
result_2 = pd.merge(result_1,
                  dataPerMonth.loc[:,['统一社会信用代码', '连续报送月']],
                  how='left', left_on='社会信用代码', right_on='统一社会信用代码',suffixes=('_1_L','_1_R'))
print(result_2.columns)
print(result_2[result_2['连续报送月_1_L'].isna()].shape,result_2[result_2['连续报送月_1_R'].isna()].shape)

func_3months = lambda x: x[1] if x[0] is np.nan else x[0]
result_2['连续报送月_final'] = result_2.loc[:,['连续报送月_1_L','连续报送月_1_R']].apply(func_3months,axis=1)

# lambda x1,x2: x1 if x2 is np.nan else x2
# 报错"<lambda>() missing 1 required positional argument: 'value2'" site:stackoverflow.com

print(result_2[result_2['连续报送月_final'].isna()].shape)



result_3 = pd.merge(result_2.drop(columns='入库率'),
                  successedUploadRatio.loc[:,['机构全称', '入库率']],
                  how='left', left_on='全称', right_on='机构全称',suffixes=('_2_L','_2_R'))
print(result_3.columns)
print(result_3[result_3['入库率'].isna()].shape)


result_4 = pd.merge(result_3,
                  successedUploadRatio.loc[:,['统一社会信用代码', '入库率']],
                  how='left', left_on='社会信用代码', right_on='统一社会信用代码',suffixes=('_3_L','_3_R'))
print(result_4.columns)
print(result_4[result_4['入库率_3_R'].isna()].shape)

func_uploadRatio = lambda x1,x2: x1 if not (x1 is np.nan) else x2
result_4['入库率_final'] = result_4.loc[:,'入库率_3_L'].combine(result_4['入库率_3_R'],func_uploadRatio)

result_4[result_4['入库率_final'].isna()].shape


result_5 = pd.merge(result_4.drop(columns='更新率'),
                  updateRatio.loc[:,['全称', '更新率']],
                  how='left', left_on='全称', right_on='全称',suffixes=('_4_L','_4_R'))
print(result_5.columns)
print(result_5[result_5['更新率'].isna()==True].shape)


result_6 = pd.merge(result_5,
                  updateRatio.loc[:,['统一社会信用代码', '更新率']],
                  how='left', left_on='社会信用代码', right_on='统一社会信用代码',suffixes=('_5_L','_5_R'))
print(result_6.columns)
print(result_6[result_6['更新率_5_R'].isna()==True].shape)


result_6['更新率_final'] = result_6.loc[:,'更新率_5_L'].combine_first(result_6.loc[:,'更新率_5_R'])
print(result_6[result_6['更新率_final'].isna()].shape)
result_6.columns
usedNames = {'机构简称', '机构全称', '', '统一社会信用代码', '管理员账号', '接入状态',
        '开通方式', '报送方式', '开通状态', '开通时间', '备注', '更新日期', 
       
       '连续报送月_final', 
       '入库率_final','更新率_final'}
result_6.drop(columns=['是否会员-旧总表', 
       '业态', '报送月份数',
       '部门', '连续报送月_1_L', '统一社会信用代码_3_L', '连续报送月_1_R',
       '机构全称', '入库率_3_L', '统一社会信用代码_3_R', '入库率_3_R',
       '更新率_5_L', '统一社会信用代码', '更新率_5_R']).columns


with pd.ExcelWriter(r'D:\...\gxpt\20190906\result.xlsx',
                       date_format='YYYY/MM/DD') as writer:
    result_6.drop(columns='业态').to_excel(writer,
                                         sheet_name='接入机构',
                                         na_rep='#N/A',
                     columns=['简称', '全称', '是否会员', '社会信用代码', '管理员账号', '接入状态', '是否持牌', '开通方式',
           '报送方式', '开通状态', '开通时间', '备注', '更新日期', '连续报送月_final',
           '入库率_final', '更新率_final'])
    result_6.drop(columns='业态').to_excel(writer,
                                         sheet_name='接入机构简版',
                                         na_rep='#N/A',
                     columns=['简称', '全称', '社会信用代码', '管理员账号', '开通方式',
           '报送方式', '开通状态', '开通时间', '连续报送月_final',
           '入库率_final', '更新率_final'])
    result_6.loc[result_6['开通状态']!='已开通',:].drop(columns='业态').to_excel(writer,
                                                                        sheet_name='未开通',
                                                                        na_rep='#N/A',
                                                                        columns=['简称', '全称', '是否会员', '社会信用代码', 
                                                                                 '管理员账号', '接入状态', '是否持牌', '开通方式',
                                                                                 '报送方式', '开通状态', '开通时间', '备注',
                                                                                 '更新日期', '连续报送月_final',
                                                                                 '入库率_final', '更新率_final'])
    # lendingHelper.to_excel(writer,
    #                      sheet_name='zzjg',
    #                      na_rep='#N/A')
    contactBook.to_excel(writer,
                         sheet_name='txl',
                         na_rep='#N/A')
    dataCatagory.to_excel(writer,
                         sheet_name='月报中机构类型',
                         na_rep='#N/A')
# ValueError: cannot reindex from a duplicate axis
# Stack Overflow上基本也是As others have said, 
# you've probably got duplicate values in your original index. To find them do this:  df[df.index.duplicated()]。

