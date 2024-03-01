# -*- coding: utf-8 -*-
"""
Created on Thu Feb 29 15:29:08 2024

@author: yangxy
"""


import pandas as pd
import numpy as np
import copy
import openpyxl as op
import os
import xlwings as xw
import datetime
import sys
os.chdir(r'D:\desktop\mycase\m1') 

app=xw.App(visible=False,add_book=False)
app.display_alerts=False
app.screen_updating=False

#第一步读取数据
dirpath=r"D:\desktop\mydatabase\usclose"
dirlist=os.listdir(dirpath)
greatlis=dict()
filenamelis=list()
for filename in dirlist:
    print(filename)
    filenamelis.append(filename)
    if os.path.isdir(os.path.join(dirpath,filename)) or filename=="ok.xlsx" or ("~$" in filename):
        continue
    wb=app.books.open(os.path.join(dirpath,filename))
    #wb[wb.sheetnames[0]].title
    greatdf=dict()
    for she in range(len(wb.sheets)):
        wb.sheets[she].name
        wb.sheets[0].used_range.last_cell.row
        wb.sheets[0].used_range.last_cell.column
        df=wb.sheets[she].range('A1').options(pd.DataFrame,index=False,expand="table").value
        #视情况看是否要对齐至季度最后一天
        greatdf[wb.sheets[she].name]=df
    greatlis[filename]=greatdf
    wb.close()
del dirlist,dirpath,df,filename,filenamelis,greatdf,she,wb

#第二步选择标的
focus=["NDX.GI"]

for i in greatlis.keys():#前提是只有一层
    for j in greatlis[i].keys():
        tmp1=greatlis[i][j]
close=tmp1.loc[:,focus]
close.index=tmp1["Date"]
del tmp1,i,j



#第三步计算MA1，MA2
ma1=5
ma2=22

for i in focus:
    tmp1=pd.DataFrame(close[i])
    tmp1["ma1"]=close[i].rolling(window=ma1, min_periods=ma1,axis=0).mean()
    tmp1["ma2"]=close[i].rolling(window=ma2, min_periods=ma2,axis=0).mean()
    tmp1=tmp1.dropna(axis=0,how="any")
    tmp1["ma1>ma2"]=(tmp1["ma1"]-tmp1["ma2"]).map(lambda x:
                                                  1 if x>0 else 0 
                                                  )
    tmp1["2ma_change"]=tmp1["ma1>ma2"]-tmp1["ma1>ma2"].shift(1)
    tmp1["close>ma1"]=(tmp1[i]-tmp1["ma1"]).map(lambda x:
                                                  1 if x>0 else 0 
                                                  )
    tmp1["close_ma_change"]=tmp1["close>ma1"]-tmp1["close>ma1"].shift(1)
    del tmp1["ma1>ma2"],tmp1["close>ma1"]
    tmp1=tmp1.dropna(axis=0,how="any")
del close,greatlis,i,ma1,ma2
df=copy.deepcopy(tmp1)
del tmp1

#第四步计算每笔交易和每日净值
trades=dict()
aum=dict()
stock=focus[0]
for i in range(len(df.index)):
    dt=df.index[i]
    if i==0:
        if df.loc[dt,"close_ma_change"]==0:
            aum[dt]={"shares":0,"aum":10000}
        elif df.loc[dt,"close_ma_change"]==1:
            aum[dt]={"shares":0,"aum":10000}
        continue
   
    dt_lag1=df.index[i-1]
    #净值的自然增长
    if aum[dt_lag1]["shares"]==0:
        aum[dt]={"shares":0,"aum":aum[dt_lag1]["aum"]}
    else:
        aum[dt]={"shares":0,
            "aum":aum[dt_lag1]["aum"]*df.loc[dt,stock]/df.loc[dt_lag1,stock]
            }
    #是否要开仓
    if df.loc[dt,"close_ma_change"]==1:
        aum[dt]["shares"]=aum[dt]["aum"]/df.loc[dt,stock]
    elif df.loc[dt,"close_ma_change"]==1:
        aum[dt]={"shares":0,"aum":10000}

            




#第二步筛选每日持仓（暂时不考虑每日净值）
#MyPortfolio、MyPortfolioDueFundsNum、MyPortfolioDueAssetRatio、MyPortfolioDueFundsChange
#MyPortfolioNoMA、MyPortfolioDueFundsNumNoMA、MyPortfolioDueAssetRatioNoMA、MyPortfolioDueFundsChangeNoMA 
my=y2.MyPortfolioDueAssetRatioNoMA(dataset,20) 
#my.init_price_and_factor()
basic_parameter=my.get_para_data()
my_portfolio=my.get_pf()
del my

#第三步计算每日净值
cn=y2.CalcuNav(my_portfolio,basic_parameter)
my_portfolio=cn.do_and_return()

#另一种思路的第二步，生成双因子











#集成总表和详表
datelist=pd.Series(list(my_portfolio.keys())).sort_values(ascending=True).reset_index(drop=True)
general_df=pd.DataFrame(columns=["日期","每日净值"])
detailed_df=pd.DataFrame()
for i in range(len(datelist)):
    trade_day=datelist[i]
    money=my_portfolio[trade_day]["价值"].sum()
    tmp1=my_portfolio[trade_day].copy()
    tmp1.index=[trade_day for j in range(len(tmp1.index))]
    detailed_df=pd.concat([detailed_df,tmp1],axis=0)
    tmp2=pd.DataFrame((trade_day,money),index=["日期","每日净值"],columns=[i]).T
    general_df=pd.concat([general_df,tmp2],axis=0)
    
#计算最大回撤
general_df["accu_max"]=[0.0 for i in range(len(general_df.index))]
for i in range(len(general_df.index)):
    tmp1=general_df.iloc[:i,1].max()
    if (tmp1==tmp1):
        general_df["accu_max"].iat[i]=general_df["每日净值"].iat[i]/tmp1-1
general_df["accu_max"].min()

#输出表格
general_df.to_excel(r'D:\desktop\fundlove\总表.xlsx',header=True,index=True)
detailed_df.to_excel(r'D:\desktop\fundlove\详表.xlsx',header=True,index=True)        





app.kill()
