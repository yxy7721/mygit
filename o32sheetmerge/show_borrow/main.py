# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 14:04:08 2024

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
from show_borrow.all_borrow_exe import get_sht
res=get_sht(app,greatdf)
del get_sht


#三、整主机构的情况
tmp1=copy.deepcopy(res["client_all"])
tmp2=copy.deepcopy(res["client_keytime"])
from show_borrow.all_borrow_exe import main_name
tmp1=main_name(tmp1)
tmp2=main_name(tmp2)
del main_name
tmp3=tmp1.groupby(["主机构"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
tmp4=tmp2.groupby(["主机构"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
tmp3.columns=["模拟利息","总成交金额(亿元)"]
tmp3["平均价格(%)"]=tmp3["模拟利息"]/100000000/tmp3["总成交金额(亿元)"]*100*365
tmp4.columns=["模拟利息","总成交金额(亿元)"]
tmp4["平均价格(%)"]=tmp4["模拟利息"]/100000000/tmp4["总成交金额(亿元)"]*100*365
tmp3=tmp3.sort_values(by="总成交金额(亿元)",ascending=False)
tmp4=tmp4.sort_values(by="总成交金额(亿元)",ascending=False)

shuchu=app.books.add()    
shuchu.sheets[0].name="all正回购情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
shuchu.sheets.add('关键时点支持情况')
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp4
shuchu.save(r'D:\desktop\所有正回购.xlsx')
del shuchu



app.quit()




























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
