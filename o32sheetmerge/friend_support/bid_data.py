# -*- coding: utf-8 -*-
"""
Created on Mon Dec 12 17:09:04 2022

@author: yangxy
"""

import pandas as pd
import numpy as np
import docx
import os
import xlwings as xw
import copy


def get_data(app,greatdf):
    
    wb=app.books.open(r"D:\desktop\全年跨月支持情况\一级投标统计-存单.xlsx")
    wb.sheets[0].name
    wb.sheets[0].used_range.last_cell.row
    wb.sheets[0].used_range.last_cell.column
    tmp1=wb.sheets[0].range('A1').options(pd.DataFrame,index=False,expand="table").value
    cd=copy.deepcopy(tmp1)
    
    wb=app.books.open(r"D:\desktop\全年跨月支持情况\一级投标统计-利率.xlsx")
    wb.sheets[0].name
    wb.sheets[0].used_range.last_cell.row
    wb.sheets[0].used_range.last_cell.column
    tmp1=wb.sheets[0].range('A1').options(pd.DataFrame,index=False,expand="table").value
    goodbond=copy.deepcopy(tmp1)
    
    wb=app.books.open(r"D:\desktop\全年跨月支持情况\一级投标统计-信用.xlsx")
    wb.sheets[0].name
    wb.sheets[0].used_range.last_cell.row
    wb.sheets[0].used_range.last_cell.column
    tmp1=wb.sheets[0].range('A1').options(pd.DataFrame,index=False,expand="table").value
    badbond=copy.deepcopy(tmp1)
    del tmp1
    
    cd=cd.rename(columns={None:"对手方"})[["对手方","合计："]]
    goodbond=goodbond.rename(columns={None:"对手方"})[["对手方","合计："]]
    badbond=badbond.rename(columns={None:"对手方"})[["对手方","合计："]]
    return goodbond,cd,badbond
    
    
def v_into(res,goodbond,cd,badbond):
    tmp3=copy.deepcopy(res["client_all"])
    tmp3["利率债中标（万元）"]=0.0
    tmp3["存单支持（万元）"]=0.0
    tmp3["信用债中标（万元）"]=0.0
    for i in range(len(cd.index)):
        tmp7=cd.iat[i,0]
        if tmp7 in tmp3.index:
            tmp3.loc[tmp7,"存单支持（万元）"]=cd.iat[i,1]
    for i in range(len(goodbond.index)):
        tmp7=goodbond.iat[i,0]
        if tmp7 in tmp3.index:
            tmp3.loc[tmp7,"利率债中标（万元）"]=goodbond.iat[i,1]
    for i in range(len(badbond.index)):
        tmp7=badbond.iat[i,0]
        if tmp7 in tmp3.index:
            tmp3.loc[tmp7,"信用债中标（万元）"]=badbond.iat[i,1]
    res["client_all"]=tmp3
    return res
























