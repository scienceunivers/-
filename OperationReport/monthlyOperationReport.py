import openpyxl,xlrd,xlwings
import pandas as pd, numpy as np
import datetime as dt

import os

# openpyxl==2.4
# xlwings==0.11

workbook_orgTable_all=openpyxl.load_workbook(r'D:\...\gxpt\cyzb\1-jgjrzb20210203.xlsx',
                                            data_only=True)
sheetNames = workbook_orgTable_all.sheetnames
print(sheetNames)
objectiveSheet = 'Sheet1'

nCols = workbook_orgTable_all[objectiveSheet].max_column
nRows = workbook_orgTable_all[objectiveSheet].max_row
print(nRows,nCols)
colNames = []
# colNames_Row1,colNames_Row2 = [],[]
workbook_orgTable_all_dict = {}
# for j in range(0,nRows):
#     if any(wworkbook_orgTable_all[objectiveSheet].iter_rows(max_col=nCols, max_row=nRows)):
#         print(j)
#         break
for jRow in workbook_orgTable_all[objectiveSheet].iter_rows(min_row=None, max_row=nRows, min_col=None, max_col=nCols, 
                                                            ):
    for jCell in jRow:
        if jCell.value:
            
            # 没用worksheet的values属性，因为values是按行返回值，
            # 但没有对偶的按列返回值的属性，如果遇到了按列返回值的需求还得重新写循环。
            print(jCell.row)
            # openpyxl计数大部分是1-base。
            jCell_value=jCell.value
            break
    if jCell.value:
        break

# 判定从哪行开始有数据，不能用workbook_orgTable_all[objectiveSheet].min_row
for iRow in workbook_orgTable_all[objectiveSheet].iter_rows(min_col=None, max_col=None, min_row=jCell.row, 
                                                                         max_row=jCell.row):
    # 迭代出表头这一行，将这一行单元格的值append到colNames中。
    for iCell in iRow:
        colNames.append(iCell.value)
        workbook_orgTable_all_dict[colNames[iCell.column-1]] = []
        # 不如xlrd直接用col_values返回list方便。
        for kColumn in workbook_orgTable_all[objectiveSheet].iter_cols(min_col=iCell.column, 
                                                                         max_col=iCell.column,
                                                                         min_row=iCell.row+1, max_row=None):
            for kCell in kColumn:
                workbook_orgTable_all_dict[colNames[iCell.column-1]].append(kCell.value)
                      
        # iCell.column-1，报错TypeError: unsupported operand type(s) for -: 'str' and 'int'，2.4版本源代码中确实把column
        # 写成了获得单元格列字母的属性，col_idx是单元格列号的属性，到2.6将column改成了列号，获得列字母的属性改成了get_column_letter。
        
# 读取工作簿的代码，也可以用Worksheet.cell(row,column).value的思路重写一下。
data = pd.DataFrame(workbook_orgTable_all_dict, columns=colNames)
originalData_jieruzongbiao=data.loc[:,:'zsjrsjc'].copy()

# openpyxl把日期列的空单元格都转化成了NaT，把文本列的空单元格转化成了NAN。



# 侦测空格：

# originalData_jieruzongbiao.where(originalData_jieruzongbiao==' ')
print(
    originalData_jieruzongbiao.loc[originalData_jieruzongbiao.isin([' ']).any(axis=1)].index,'\n',
    originalData_jieruzongbiao.loc[:,originalData_jieruzongbiao.isin([' ']).any(axis=0)].columns
)
originalData_jieruzongbiao.replace(to_replace=r'^[ \t\n\r\f\v]+|[ \t\n\r\f\v]+$',value=np.nan,regex=True,inplace=True)
# 如果to_replace, value参数中任何一个采取regex形式，需令regex=True。如果value=None，则不替换。虽然None与np.nan
# 在df中涉及df数值运算时被同等对待，但在字符串操作比如替换中，二者又不一样。

originalData_jieruzongbiao.fillna({'xysjc':pd.NaT,'pxsjc':pd.NaT,'jrcssjc':pd.NaT,'zsjrsjc':pd.NaT}
                                          ,inplace=True)
print(
    originalData_jieruzongbiao.loc[originalData_jieruzongbiao.isin([' ']).any(axis=1)].index,'\n',
    originalData_jieruzongbiao.loc[:,originalData_jieruzongbiao.isin([' ']).any(axis=0)].columns
)
originalData_jieruzongbiao.loc[:,['xysjc','pxsjc','jrcssjc','zsjrsjc']].dtypes 
# 验证各类时间戳类型。
# originalData_jieruzongbiao['正式接入时间戳'].apply(lambda x: 0 if isinstance(x,(dt.datetime,pd._libs.tslibs.nattype.NaTType)) else 1)
# 后续如有其它脏数据，再改regex。



# 后续完善数据质量检测代码：
# excel把空格也当做筛选中的空白。
# 虽然python认为None ==None为真，但在pandas里，None按np.nan对待，不建议在pandas中用None ==None。



# 正式接入机构。
officiallyParticiOrg_jieruzongbiao=originalData_jieruzongbiao.loc[
    (originalData_jieruzongbiao['hzxy']!='4协议解除')
        &~(originalData_jieruzongbiao['zsjr'].isna())
           &(originalData_jieruzongbiao['zsjrsjc']<=dt.datetime(2021,1,31)),:
    ]
        
officiallyParticiOrg_jieruzongbiao_danweiquancheng=officiallyParticiOrg_jieruzongbiao.loc[:,'单位全称'
    ]
print(
    '截至月末正式接入报数机构{}家({:%Y%m%d})'.format
    (
   officiallyParticiOrg_jieruzongbiao.shape[0],dt.datetime.today()
    )
)
# openpyxl\cell.py, openpyxl\styles\numbers.py源代码，用其中方法
# openpyxl.styles.numbers.is_date_format(workbook_orgTable_all[objectiveSheet].cell(779,21).value)去判定空时间戳类型。
officiallyParticiOrg_jieruzongbiao_20210203=officiallyParticiOrg_jieruzongbiao.copy()



# 当前接入机构情况汇总。
originalData_jierujigouhuizongbiao = pd.read_excel(
    r'...\gxpt\cyzb\202102202055gxptdqjrjgqkhz.xlsx',
    sheet_name=0,header=0, names=None, index_col=None, 
    usecols=None, squeeze=False, dtype=None, engine=None, 
    converters=None, true_values=None, false_values=None, skiprows=None, nrows=None, 
    na_values=None, keep_default_na=True, verbose=False, 
    parse_dates=False, date_parser=None, thousands=None, comment=None, 
    skipfooter=0, convert_float=True, mangle_dupe_cols=True
)
originalData_jiekoujierubiao = pd.read_excel(
    r'...\gxpt\cyzb\6-T+1jkbsjdb20210201.xlsx',
    sheet_name=0,header=0, names=None, index_col=None, 
    usecols=None, squeeze=False, dtype=None, engine=None, 
    converters=None, true_values=None, false_values=None, skiprows=None, nrows=None, 
    na_values=None, keep_default_na=True, verbose=False, 
    parse_dates=False, date_parser=None, thousands=None, comment=None, 
    skipfooter=0, convert_float=True, mangle_dupe_cols=True
)

# 1224后续要读取两个月份的接口接入进度表，还是把更新日期列统一成datetime64更方便。通过如下代码：
# originalData_jiekoujierubiao.loc[
#     originalData_jiekoujierubiao['更新日期'].str.contains(r'^[ \t\n\r\f\v]+|[ \t\n\r\f\v]+$', regex=True).notna(),:]
# 检测出某个单元格是string且不含whitespace，是False，非string的元素都返回NaN，这应该就是字符了，输出发现果然是备注列填到了更新日期列。
officiallyParticiOrg_jiekoujierubiao=originalData_jiekoujierubiao.loc[
    (originalData_jiekoujierubiao['jzqk']=='已完成')
    &(
        (originalData_jiekoujierubiao['更新日期'].isna())
        |(originalData_jiekoujierubiao['更新日期']<=dt.datetime(2021,1,31))
    ),:
]
# pd.NaT>dt.datetime(2020,11,30)，无论>, <,都返回False，与nan类似。
print(
    '截至月末接口接入报数机构{}家，网页接入机构{}家({:%Y%m%d})'.format
    (
        officiallyParticiOrg_jiekoujierubiao.shape[0],
        officiallyParticiOrg_jieruzongbiao.shape[0]-officiallyParticiOrg_jiekoujierubiao.shape[0],
        dt.datetime.today()
    )
)

officiallyParticiOrg_jiekoujierubiao_20210201=officiallyParticiOrg_jiekoujierubiao.copy()
