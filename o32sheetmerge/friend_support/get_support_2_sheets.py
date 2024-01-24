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


def get_sht(app,greatdf):
    tmp1=(greatdf[ 
            (greatdf["业务分类"]=="银行间业务") ]        
            ).copy()
    tmp1=(tmp1[ 
            (tmp1["委托方向"]=="融资回购／拆入") |
            (tmp1["委托方向"]=="融资回购／拆入")]        
            ).copy()
    tmp2=tmp1["基金名称"].map(lambda x: False if x.find("号")!=-1 else True)
    tmp1=tmp1[tmp2]
    del tmp2
    
    wb=app.books.open(r"D:\desktop\12月合作机构支持情况\基础数据.xlsx")
    wb.sheets[0].name
    wb.sheets[0].used_range.last_cell.row
    wb.sheets[0].used_range.last_cell.column
    friendcop=wb.sheets[0].range('A1').options(pd.DataFrame,index=False,expand="table").value
    friendcop=list(friendcop["银行"])
    tmp2=tmp1["交易对手"].map(lambda x:True if x in friendcop else False)
    tmp1=copy.deepcopy(tmp1[tmp2])
    tmp1=tmp1.reset_index(drop=True)
    del tmp2
    
    #分银行的类别，暂时用不上
    tmp2=["other" for i in range(len(friendcop))]
    belongto = dict(zip(friendcop, tmp2))
    soebanklis="招商银行、浦发银行、中信银行、光大银行、华夏银行、民生银行、广发银行、兴业银行、平安银行、浙商银行、恒丰银行、渤海银行"
    soebanklis=soebanklis.split("、")
    biglis="工商银行、农业银行、中国银行、建设银行、交通银行、邮储银行、进出口行"
    biglis=biglis.split("、")
    for i in belongto.keys():
        print(i)
        if i[-1]=="行" or i[-1]=="社":
            if i in soebanklis:
                belongto[i]="国股行"
            elif i in biglis:
                belongto[i]="超大行"
            elif i.find("农")!=-1:
                belongto[i]="农商行"
            else:
                belongto[i]="城商行"
        else:
            belongto[i]="券商自营" 
    del biglis,i,soebanklis,tmp2
    
    tmp1["日期"]=pd.to_datetime(tmp1["日期"])
    tmp1["月份"]=tmp1["日期"].map(lambda x:x.month)
    tmp1["指令价格(主币种)"]=tmp1["指令价格(主币种)"].astype(float)
    tmp1["到期清算金额"]=tmp1["到期清算金额"].map(lambda x:float(x.replace(",","")))
    tmp1["实际利息率"]=(tmp1["到期清算金额"]/tmp1["当日成交金额"]-1)*365
    tmp1["回购天数"]=(tmp1["实际利息率"]/tmp1["指令价格(主币种)"]*100).round(0)
    tmp1["交易对手归类"]=tmp1["交易对手"].map(lambda x:belongto[x])
    tmp1["day_shift"]=tmp1["回购天数"].map(lambda x:x*pd.offsets.Day())
    tmp1["到期月份"]=(tmp1["day_shift"]+tmp1["日期"]).map(lambda x:x.month)
    tmp1["到期月份"]=(tmp1["到期月份"]-tmp1["月份"]).map(lambda x:True if x!=0 else False)
    tmp1["模拟利息"]=tmp1["当日成交金额"]*tmp1["指令价格(主币种)"]/100/365
    tmp2=copy.deepcopy(tmp1[tmp1["到期月份"]])
    
    #tmp1是全，tmp2是跨，下面这是对对手方为维度
    tmp3=tmp1.groupby(["交易对手"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
    tmp3["平均价格"]=tmp3["模拟利息"]/tmp3["当日成交金额"]*100*365
    tmp3["当日成交金额"]=tmp3["当日成交金额"]/100000000
    tmp4=tmp2.groupby(["交易对手"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
    tmp4["平均价格"]=tmp4["模拟利息"]/tmp4["当日成交金额"]*100*365
    tmp4["当日成交金额"]=tmp4["当日成交金额"]/100000000
    '''
    shuchu=app.books.add()    
    shuchu.sheets[0].name="合作机构这个月正回购情况"
    shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
    shuchu.sheets.add('合作机构跨月支持情况')
    shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp4
    shuchu.save(r'D:\desktop\交易对手维度.xlsx')
    del shuchu
    '''
    
    res={"client_all":copy.deepcopy(tmp3),
         "client_keytime":copy.deepcopy(tmp4)    
        }
    
    #tmp1是全，tmp2是跨，下面这是对我方产品为维度
    wb=app.books.open(r"D:\desktop\12月合作机构支持情况\产品类型基础数据.xlsx")
    wb.sheets[0].name
    wb.sheets[0].used_range.last_cell.row
    wb.sheets[0].used_range.last_cell.column
    fundtype=wb.sheets[0].range('A1').options(pd.DataFrame,index=False,expand="table").value
    fundtype=dict(zip(fundtype.iloc[:,0],fundtype.iloc[:,1]))
    
    tmp3=tmp1.groupby(["基金名称"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
    tmp3["平均价格"]=tmp3["模拟利息"]/tmp3["当日成交金额"]*100*360
    tmp3["当日成交金额"]=tmp3["当日成交金额"]/100000000
    tmp4=tmp2.groupby(["基金名称"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
    tmp4["平均价格"]=tmp4["模拟利息"]/tmp4["当日成交金额"]*100*360
    tmp4["当日成交金额"]=tmp4["当日成交金额"]/100000000
    tmp3["类型"]=tmp3.index.map(lambda x: fundtype[x])
    tmp4["类型"]=tmp4.index.map(lambda x: fundtype[x])
    tmp3=tmp3[[ '类型','模拟利息', '当日成交金额', '平均价格']]
    tmp4=tmp4[[ '类型','模拟利息', '当日成交金额', '平均价格']]
    '''
    shuchu=app.books.add()    
    shuchu.sheets[0].name="合作机构这个月正回购情况"
    shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
    shuchu.sheets.add('合作机构跨月支持情况')
    shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp4
    shuchu.save(r'D:\desktop\产品维度.xlsx')
    del shuchu
    '''
    res["me_all"]=copy.deepcopy(tmp3)
    res["me_keytime"]=copy.deepcopy(tmp4)    
    return res
    
    




























