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
from o32_read_data import read_data_from_temp
greatdf=read_data_from_temp(app,dirpath)
#至此以上，读取已完毕



#二、跨和正常时节的支持的
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
tmp1["模拟利息"]=tmp1["当日成交金额"]*tmp1["指令价格(主币种)"]/100/360
tmp2=copy.deepcopy(tmp1[tmp1["到期月份"]])

tmp3=tmp1.groupby(["交易对手"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
tmp3["平均价格"]=tmp3["模拟利息"]/tmp3["当日成交金额"]*100*360
tmp3["当日成交金额"]=tmp3["当日成交金额"]/100000000
tmp4=tmp2.groupby(["交易对手"]).agg({"模拟利息":"sum","当日成交金额":"sum"})
tmp4["平均价格"]=tmp4["模拟利息"]/tmp4["当日成交金额"]*100*360
tmp4["当日成交金额"]=tmp4["当日成交金额"]/100000000

shuchu=app.books.add()    
shuchu.sheets[0].name="合作机构这个月正回购情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
shuchu.sheets.add('合作机构跨月支持情况')
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp4
shuchu.save(r'D:\desktop\交易对手维度.xlsx')
del shuchu

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

shuchu=app.books.add()    
shuchu.sheets[0].name="合作机构这个月正回购情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
shuchu.sheets.add('合作机构跨月支持情况')
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp4
shuchu.save(r'D:\desktop\产品维度.xlsx')
del shuchu

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

#跨不看合作机构
tmp1=(greatdf[ 
        (greatdf["业务分类"]=="银行间业务") ]        
        ).copy()
tmp1=(tmp1[ 
        (tmp1["委托方向"]=="融资回购／拆入") |
        (tmp1["委托方向"]=="融券回购／拆出")]        
        ).copy()

tmp1["日期"]=pd.to_datetime(tmp1["日期"])
tmp1["月份"]=tmp1["日期"].map(lambda x:x.month)
tmp1["指令价格(主币种)"]=tmp1["指令价格(主币种)"].astype(float)
tmp1["到期清算金额"]=tmp1["到期清算金额"].map(lambda x:float(x.replace(",","")))
tmp1["实际利息率"]=(tmp1["到期清算金额"]/tmp1["当日成交金额"]-1)*365
tmp1["回购天数"]=(tmp1["实际利息率"]/tmp1["指令价格(主币种)"]*100).round(0)
tmp1["day_shift"]=tmp1["回购天数"].map(lambda x:x*pd.offsets.Day())
tmp1["到期月份"]=(tmp1["day_shift"]+tmp1["日期"]).map(lambda x:x.month)
tmp1["到期月份"]=(tmp1["到期月份"]-tmp1["月份"]).map(lambda x:True if x!=0 else False)
tmp1["模拟利息"]=tmp1["当日成交金额"]*tmp1["指令价格(主币种)"]/100/360
tmp1=tmp1.reset_index(drop=True)
tmp2=copy.deepcopy(tmp1[tmp1["到期月份"]])
tmp3=tmp2[tmp2["证券代码"]=="R004"].index
tmp2.loc[tmp3,"证券名称"]="R001"

tmp1.columns
tmp3=tmp2.groupby(["证券名称","委托方向","月份"]).agg({"模拟利息":"sum","当日成交金额":"sum","价格模式":"count"})
tmp3["平均价格"]=tmp3["模拟利息"]/tmp3["当日成交金额"]*100*360
tmp3["当日成交金额"]=tmp3["当日成交金额"]/100000000

shuchu=app.books.add()    
shuchu.sheets[0].name="合作机构全年正回购情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
shuchu.save(r'D:\desktop\跨年正逆回购工作量情况.xlsx')
del shuchu,tmp3






#分交易员使用
tmp1=greatdf[(greatdf['交易员']=="曹越") | (greatdf["交易员"]=="陈鑫") | (greatdf["交易员"]=="杨新宇")].copy()
greatdf=tmp1.copy()
del tmp1

shuchu=app.books.add()    
shuchu.sheets[0].name="指令全"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=greatdf
shuchu.save(r'D:\desktop\o32sheetmerge\merged_all.xlsx')
del shuchu

tmp1=greatdf[greatdf["交易员"]=="曹越"].copy()
shuchu=app.books.add()    
shuchu.sheets[0].name="场内指令CY"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp1
shuchu.save(r'D:\desktop\o32sheetmerge\calcu_number.xlsx')
del shuchu,tmp1

tmp1=greatdf[greatdf["交易员"]=="陈鑫"].copy()
shuchu=app.books.add()    
shuchu.sheets[0].name="场内指令CX"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp1
shuchu.save(r'D:\desktop\o32sheetmerge\merged_cx.xlsx')
del shuchu,tmp1

tmp1=greatdf[greatdf["交易员"]=="杨新宇"].copy()
shuchu=app.books.add()    
shuchu.sheets[0].name="场内指令YXY"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp1
shuchu.save(r'D:\desktop\o32sheetmerge\merged_yxy.xlsx')
del shuchu,tmp1

#分交易员使用（整个循环）
namelist=["杨新宇","胡林峰","罗馨","杜小宇","刘坤虹","吴靉倩","牛丽珠"
          ,"杨雪"]
for name in namelist:
    tmp1=greatdf[greatdf["执行人"]==name].copy()
    shuchu=app.books.add()    
    shuchu.sheets[0].name="指令"
    shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp1
    shuchu.save(r"D:\desktop\\" + name+".xlsx")
    



app.quit()

























