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

app=xw.App(visible=False,add_book=False)
app.display_alerts=False
app.screen_updating=False
os.chdir(r'D:\desktop\mycase\o32sheetmerge')


#一、先读取fundhold

dirpath=r"D:\desktop\指令全存档\temp"
from read_data.o32_read_data import read_data_from_temp
greatdf=read_data_from_temp(app,dirpath)
del dirpath,read_data_from_temp
#至此以上，读取已完毕



#二、生成表，分对方和我方两个维度
from friend_support.get_support_2_sheets import get_sht
res=get_sht(app,greatdf)
del get_sht


#三、导入投标情况





#下面是生成文字部分
    tmp3=tmp3.reset_index(drop=False)
    tmp4=tmp4.reset_index(drop=False)
    tmp5=tmp3.groupby(["类型"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
    tmp6=tmp4.groupby(["类型"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
    
    tmp7=list()
    for i in list(tmp5.index):
        tmp7.append(i+"总金额"+"%.2f" % tmp5.loc[i,"当日成交金额"]+"亿元")
    ",".join(tmp7)
    
    tmp7=list()
    for i in list(tmp6.index):
        tmp7.append(i+"总金额"+"%.2f" % tmp6.loc[i,"当日成交金额"]+"亿元")
    ",".join(tmp7)
    
    
    del tmp1,tmp2,friendcop


app.quit()

























