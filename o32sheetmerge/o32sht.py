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


#一、先读取fundhold

#os.chdir(r"D:\desktop\index_research\datamix\data")
dirpath=r"D:\desktop\指令全存档\temp"
dirlist=os.listdir(dirpath)

'''
hang=len(wb.sheets[0].range('A1').current_region.rows)
lie=len(wb.sheets[0].range('A1').current_region.columns)
df=wb.sheets[0].range((1,1),(hang,lie)).options(pd.DataFrame,index=False).value
'''

greatlis=dict()
for filename in dirlist:
    print(filename)
    #break
    wb=app.books.open(os.path.join(dirpath,filename))
    #wb[wb.sheetnames[0]].title
    greatdf=dict()
    she=0
    wb.sheets[she].name
    wb.sheets[0].used_range.last_cell.row
    wb.sheets[0].used_range.last_cell.column
    df=wb.sheets[she].range('A1').options(pd.DataFrame,index=False,expand="table").value
    greatdf[wb.sheets[she].name]=df
    greatlis[filename]=greatdf
    wb.close()
del dirlist,dirpath,filename,greatdf,she,wb,df

greatdf=pd.DataFrame()
for i in greatlis.keys():
    for j in greatlis[i].keys():
        #break
        tmp1=greatlis[i][j].copy()
        print(len(tmp1.columns))
        greatdf=pd.concat([greatdf,tmp1.iloc[:-1,:]],axis=0)
del greatlis,i,j,tmp1

shuchu=app.books.add()    
shuchu.sheets[0].name="指令全"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=greatdf
shuchu.save(r'D:\desktop\2024Q1')
del shuchu

#至此以上，读取已完毕

#仅找出银行间的
greatdf.columns
tmp1=(greatdf[ 
        (greatdf["业务分类"]=="银行间业务") ]        
        ).copy()
tmp2=tmp1.head()
tmp2=tmp1["交易对手"].unique()

shuchu=app.books.add()    
shuchu.sheets[0].name="指令全"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp1
shuchu.save(r'D:\desktop\temp.xlsx')
del shuchu

#仅找出委托的户
"浚源1号" in greatdf["基金名称"].unique()
tmp1=(greatdf[
        (greatdf['基金名称']=="财硕1号") | 
        (greatdf['基金名称']=="财硕2号") | 
        (greatdf['基金名称']=="财硕3号") | 
        (greatdf['基金名称']=="财昱1号") |
        (greatdf['基金名称']=="财昱2号") |
        (greatdf['基金名称']=="浚源1号") |
        (greatdf['基金名称']=="财鑫233号") |
        (greatdf['基金名称']=="财鑫235号") |
        (greatdf['基金名称']=="财睿1号")]
        ).copy()
tmp1["委托方向"].unique()
tmp1.columns
tmp1["交易市场"]=tmp1["交易市场"].map(lambda x:"交易所" if x.find("交所")!=-1 else x)
tmp1["委托方向"]=tmp1["委托方向"].map(lambda x:"逆回购" if x.find("融券回购／拆出")!=-1 else x)
tmp1["委托方向"]=tmp1["委托方向"].map(lambda x:"正回购" if x.find("融资回购／拆入")!=-1 else x)
tmp1["委托方向"]=tmp1["委托方向"].map(lambda x:"场内正回购" if x.find("融资回购")!=-1 else x)
tmp1["委托方向"]=tmp1["委托方向"].map(lambda x:"场内逆回购" if x.find("融券回购")!=-1 else x)
tmp1["委托方向"]=tmp1["委托方向"].map(lambda x:"现券" if x.find("债券")!=-1 else x)



tmp2=tmp1.groupby(["交易市场","委托方向"]).agg({"交易市场":"count","当日委托金额":"sum"})
tmp3=tmp1.groupby(["交易对手","委托方向"]).agg({"交易对手":"count","当日委托金额":"sum"})
tmp4=tmp1.groupby(["基金名称","交易市场","委托方向"]).agg({"交易对手":"count","当日委托金额":"sum"})


tmp2.columns=["市场","当日委托"]
tmp2=tmp2.reset_index(drop=False)

tmp2["交易市场"]=tmp2["交易市场"].astype("category")
tmp2["委托方向"]=tmp2["委托方向"].astype("category")
tmp2['交易市场'].cat.reorder_categories(["银行间","交易所"], inplace=True)
tmp2['委托方向'].cat.reorder_categories(["正回购","正回购续做","逆回购","现券","场内正回购","场内逆回购","分销买入"], inplace=True)
tmp2['当日委托']=tmp2['当日委托']/100000000
tmp2=tmp2.sort_values(['交易市场',"委托方向"],ascending=[1,1])
tmp2=tmp2.reset_index(drop=1)


shuchu=app.books.add()    
shuchu.sheets[0].name="1"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp2
#shuchu.sheets.add('2')
#shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
shuchu.save(r'D:\desktop\temp.xlsx')
del shuchu














#计算每日groupby
df=copy.deepcopy(greatdf)
df.columns
tmp1=df.groupby("日期").agg({"日期":"count"})
tmp1.index.name="index"
tmp1.sum()
tmp1.sort_values(by="日期",ascending=0)
tmp1.mean()

tmp2=(greatdf[
        (greatdf['业务分类']=="交易所大宗交易")]
        ).copy()

tmp1=(greatdf[
        (greatdf['业务分类']=="交易所业务")]).copy()

tmp1["委托方向"].unique()
tmp1["当日委托金额"].sum()/100000000


set(greatdf['业务分类'])
set(greatdf['基金名称'])

#====================================开始分模块
greatdf.columns
tmp1=(greatdf[
        (greatdf['业务分类']=="上交所固定收益平台") | 
        (greatdf["业务分类"]=="交易所债券借贷业务") | 
        (greatdf["业务分类"]=="交易所大宗交易") | 
        (greatdf["业务分类"]=="银行间业务") |
        (greatdf["业务分类"]=="银行间买断式回购") |
        (greatdf["业务分类"]=="银行间协议转让")]
        ).copy()
tmp1["当日成交金额"].sum()/100000000
#tmp1["部门"]=tmp1['基金名称'].apply(lambda x:depart[depart["基金名称"]==x]["部门"].iat[0])
#tmp1就是固收的

#分固收公募、固收私募大类=============
len(tmp1["日期"].unique())
tmp1["部门"]=tmp1["基金名称"].map(lambda x:"私募" if x.find("号")!=-1 else x)
tmp1["部门"].unique()
tmp1["部门"]=tmp1["部门"].map(lambda x:"私募" if x.find("季季红")!=-1 else "公募")

#找出债券交易（银行间+上固收)====================================
greatdf.columns
tmp1=(greatdf[
        (greatdf['业务分类']=="银行间业务") | 
        (greatdf["业务分类"]=="上交所固定收益平台") | 
        (greatdf["业务分类"]=="交易所大宗交易") ]
        ).copy()
tmp1=(tmp1[
        (tmp1['委托方向']=="债券买入") | 
        (tmp1["委托方向"]=="债券卖出") ]
        ).copy()

shuchu=app.books.add()    
shuchu.sheets[0].name="1"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp1
#shuchu.sheets.add('2')
#shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=tmp3
shuchu.save(r'D:\desktop\三年多的债券交易总.xlsx')
del shuchu







#权益类数据精细处理============================================
tmp1["Y品种"]=tmp1["委托方向"].map(lambda x:"场内回购" if ("回购" in x) else x)
tmp1['Y品种'].unique()#just for check
tmp2=tmp1[tmp1['Y品种']=="申购"]#just for check

tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"场内债券" if ("债券" in x) else x)
tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"提交转回质押" if ("质押" in x) else x)

tmp1["tmp"]=(tmp1["业务分类"]=="网下申购") & (tmp1["Y品种"]=="申购")
tmp1.loc[((tmp1["业务分类"]=="网下申购") & (tmp1["tmp"])),"Y品种"]="网下申购"
tmp1["tmp"]=(tmp1["业务分类"]=="交易所业务") & (tmp1["Y品种"]=="申购")
tmp1.loc[((tmp1["业务分类"]=="交易所业务") & (tmp1["tmp"])),"Y品种"]="网上申购"

tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"期货交易" if ("仓" in x) else x)
tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"融资融券交易" if ("还" in x)or(x=="融资买入") else x)
tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"指定及撤销" if ("指定" in x) else x)
tmp1.loc[(tmp1["业务分类"]=="融资融券信用业务"),"Y品种"]="融资融券交易"

tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"股票买入" if ("买入" in x) else x)
tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"股票卖出" if ("卖出" in x) else x)
tmp1["Y品种"]=tmp1["Y品种"].map(lambda x:"股票混合" if ("混合" in x) else x)
del tmp2

tmp1["优于与否"]=tmp1["Y品种"].map(lambda x:x if ("股票" in x) else "不适用")

tmp1["tmp1"]=(tmp1["当日成交均价(主币种)"]>tmp1['市场有效均价(主币种)']) & (tmp1["优于与否"]=="股票卖出")
tmp1["tmp2"]=(tmp1["当日成交均价(主币种)"]<tmp1['市场有效均价(主币种)']) & (tmp1["优于与否"]=="股票买入")
tmp1["tmp"]=tmp1["tmp1"] | tmp1["tmp2"]
tmp1.loc[(tmp1["tmp"]) 
         & 
         (tmp1["优于与否"]!="不适用") 
         & 
         (tmp1["优于与否"]!="股票混合"),
         "优于与否"]="优于"
tmp1.loc[(tmp1["tmp"]==False) 
         & 
         (tmp1["优于与否"]!="不适用") 
         & 
         (tmp1["优于与否"]!="股票混合"),
         "优于与否"]="不优于"
tmp1["Y品种"].unique()

tmp2=copy.deepcopy(tmp1[(tmp1["Y品种"]=="股票买入") |
          (tmp1["Y品种"]=="股票卖出") |
          (tmp1["Y品种"]=="股票混合") ] )
tmp2["优于与否"].unique()
tmp2.columns
tmp2["日期"]=pd.to_datetime(tmp2["日期"])
tmp2["月份"]=tmp2["日期"].map(lambda x:x.strftime("%m"))

tmp3=tmp2.groupby("月份").agg({"月份":"count"})
tmp3=tmp2.groupby("优于与否").agg({"优于与否":"count"})
tmp3.loc["优于"]


shuchu=app.books.add()    
shuchu.sheets[0].name="指令全"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp2
shuchu.save(r'D:\desktop\20231229股票指令')
del shuchu

shuchu=app.books.add()    
shuchu.sheets[0].name="指令全"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp1
shuchu.save(r'D:\desktop\20231229权益指令')
del shuchu




#银行间计算利差======================================
tmp1=(greatdf[
        greatdf['业务分类']=="银行间业务" ]
        ).copy()
tmp1=(tmp1[
        (tmp1['委托方向']=="融券回购／拆出" ) |
        (tmp1['委托方向']=="融资回购／拆入" ) ]
        ).copy()
tmp1["基金名称"].unique()
tmp1["部门"]=tmp1["基金名称"].map(lambda x:"私募" if (x.find("号")!=-1) else x)
tmp1["部门"]=tmp1["部门"].map(lambda x:"私募" if (x.find("季季红")!=-1)or(x=="私募") else "公募")
tmp1["部门"].unique()

tmp1['委托方向']=tmp1['委托方向'].replace("融券回购／拆出","逆回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("逆回购续做","逆回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("融资回购／拆入","正回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("正回购续做","正回购").copy()
set(tmp1['委托方向'])
set(tmp1['回购天数'])
tmp1.columns
tmp1=tmp1.reset_index(drop=True)
tmp2=tmp1[["日期","组合名称","委托方向","指令价格(主币种)","回购天数","当日成交金额","部门"]].copy()
tmp2["日期"]=pd.to_datetime(tmp2["日期"])
tmp2["月份"]=tmp2["日期"].map(lambda x:x.month)
tmp2["指令价格(主币种)"]=tmp2["指令价格(主币种)"].astype(float)
tmp2["资金占用"]=tmp2["当日成交金额"]*tmp2["回购天数"]
tmp2["模拟利息"]=tmp2["资金占用"]*tmp2["指令价格(主币种)"]/100/360

tmp2.columns
tmp3=tmp2.groupby(["委托方向"]).agg({"模拟利息":"sum","资金占用":"sum"})
tmp3["平均价格"]=tmp3["模拟利息"]/tmp3["资金占用"]*100*360
tmp3=tmp3.reset_index(drop=False)
tmp4=tmp2.groupby(["组合名称","委托方向","月份"]).agg({"模拟利息":"sum","资金占用":"sum"})
tmp4["平均价格"]=tmp4["模拟利息"]/tmp4["资金占用"]*100*360
tmp4=tmp4.reset_index(drop=False)

shuchu=app.books.add()    
shuchu.sheets[0].name="分月明细情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp4
shuchu.sheets.add('银行间正逆回购利差')
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp3
shuchu.save(r'D:\desktop\银行间正逆回购利差.xlsx')
del shuchu,tmp3,tmp4

tmp5=tmp2.groupby(["部门","组合名称","委托方向","月份"]).agg({"模拟利息":"sum","资金占用":"sum","月份":"count"})
tmp5["平均价格"]=tmp5["模拟利息"]/tmp5["资金占用"]*100*360
tmp5=tmp5.rename(columns={"月份":"笔数"})
tmp5=tmp5.reset_index(drop=False)

shuchu=app.books.add()    
shuchu.sheets[0].name="分月明细情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp5
shuchu.save(r'D:\desktop\银行间分公私募正逆回购利差.xlsx')
del shuchu







#协回模式,计算利差=======================
tmp1=(greatdf[
        (greatdf['业务分类']=="上交所固定收益平台") | 
        (greatdf["业务分类"]=="交易所大宗交易")    ]
        ).copy()
tmp1=(tmp1[
        (tmp1['委托方向']=="正回购") | 
        (tmp1['委托方向']=="正回购续做") |
        (tmp1['委托方向']=="逆回购") |
        (tmp1['委托方向']=="逆回购续做")]
        ).copy()
tmp1["基金名称"].unique()
tmp1["部门"]=tmp1["基金名称"].map(lambda x:"私募" if (x.find("号")!=-1) else x)
tmp1["部门"]=tmp1["部门"].map(lambda x:"私募" if (x.find("季季红")!=-1)or(x=="私募") else "公募")
tmp1["部门"].unique()

tmp1['委托方向']=tmp1['委托方向'].replace("融券回购／拆出","逆回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("逆回购续做","逆回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("融资回购／拆入","正回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("正回购续做","正回购").copy()
tmp1=tmp1.reset_index(drop=True)

tmp2=tmp1[["日期","组合名称","委托方向","指令价格(主币种)","回购天数","当日成交金额","部门"]].copy()
tmp2["日期"]=pd.to_datetime(tmp2["日期"])
tmp2["月份"]=tmp2["日期"].map(lambda x:x.month)
tmp2["指令价格(主币种)"]=tmp2["指令价格(主币种)"].astype(float)
tmp2["资金占用"]=tmp2["当日成交金额"]*tmp2["回购天数"]
tmp2["模拟利息"]=tmp2["资金占用"]*tmp2["指令价格(主币种)"]/100/360

tmp2.columns
tmp3=tmp2.groupby(["组合名称","委托方向"]).agg({"模拟利息":"sum","资金占用":"sum"})
tmp3["平均价格"]=tmp3["模拟利息"]/tmp3["资金占用"]*100*360
tmp3=tmp3.reset_index(drop=False)
tmp4=tmp2.groupby(["组合名称","委托方向","月份"]).agg({"模拟利息":"sum","资金占用":"sum"})
tmp4["平均价格"]=tmp4["模拟利息"]/tmp4["资金占用"]*100*360
tmp4=tmp4.reset_index(drop=False)

shuchu=app.books.add()    
shuchu.sheets[0].name="分月明细情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp4
shuchu.sheets.add('协回正逆回购利差')
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp3
shuchu.save(r'D:\desktop\协回正逆回购利差.xlsx')
del shuchu,tmp3,tmp4

tmp5=tmp2.groupby(["部门","组合名称","委托方向","月份"]).agg({"模拟利息":"sum","资金占用":"sum","月份":"count"})
tmp5["平均价格"]=tmp5["模拟利息"]/tmp5["资金占用"]*100*360
tmp5=tmp5.rename(columns={"月份":"笔数"})
tmp5=tmp5.reset_index(drop=False)

shuchu=app.books.add()    
shuchu.sheets[0].name="分月明细情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp5
shuchu.save(r'D:\desktop\协回分公私募正逆回购利差.xlsx')
del shuchu


#计算回购现券笔数和金额
tmp1=(greatdf[
        (greatdf['业务分类']=="上交所固定收益平台") | 
        (greatdf["业务分类"]=="交易所大宗交易") | 
        (greatdf["业务分类"]=="银行间业务") ]
        ).copy()
tmp1['委托方向'].unique()
tmp1=(tmp1[
        (tmp1['委托方向']=="正回购") | 
        (tmp1['委托方向']=="正回购续做") |
        (tmp1['委托方向']=="逆回购") |
        (tmp1['委托方向']=="逆回购续做") |
        (tmp1['委托方向']=="融券回购／拆出" ) |
        (tmp1['委托方向']=="债券卖出") |
        (tmp1['委托方向']=="债券买入") |
        (tmp1['委托方向']=="融资回购／拆入" )]
        ).copy()
tmp1['委托方向']=tmp1['委托方向'].replace("融券回购／拆出","逆回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("逆回购续做","逆回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("融资回购／拆入","正回购").copy()
tmp1['委托方向']=tmp1['委托方向'].replace("正回购续做","正回购").copy()
tmp1['交易市场']=tmp1['交易市场'].replace("上交所A","交易所").copy()
tmp1['交易市场']=tmp1['交易市场'].replace("深交所A","交易所").copy()
tmp1=tmp1.reset_index(drop=True)

tmp2=tmp1[["日期","组合名称","交易市场","委托方向","指令价格(主币种)","回购天数","当日成交金额"]].copy()
tmp2["日期"]=pd.to_datetime(tmp2["日期"])
tmp2["月份"]=tmp2["日期"].map(lambda x:x.month)
tmp3=tmp2.groupby(["组合名称","交易市场","委托方向"]).agg({"指令价格(主币种)":"count","当日成交金额":"sum"})
tmp3=tmp3.reset_index(drop=False)
tmp3.columns=['组合名称', "交易市场",'委托方向', '笔数', '成交金额']
tmp4=tmp2.groupby(["组合名称","交易市场","委托方向","月份"]).agg({"指令价格(主币种)":"count","当日成交金额":"sum"})
tmp4=tmp4.reset_index(drop=False)
tmp4.columns=['组合名称', "交易市场",'委托方向', "月份",'笔数', '成交金额']

shuchu=app.books.add()    
shuchu.sheets[0].name="分月明细情况"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp4
shuchu.sheets.add('各类笔数金额')
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=tmp3
shuchu.save(r'D:\desktop\各账户回购现券笔数金额.xlsx')
del shuchu,tmp2,tmp3,tmp4







#看权益优于的，建议用上面那个=====================
for i in range(len(tmp1.index)):
    if tmp1['委托方向'][i]=='混合':
        tmp1['优于'].iat[i]="yes"
    elif tmp1['委托方向'][i]=='买入' and tmp1['当日成交均价(主币种)'][i]<tmp1['市场有效均价(主币种)'][i]:
        tmp1['优于'].iat[i]="yes"
    elif tmp1['委托方向'][i]=='卖出' and tmp1['当日成交均价(主币种)'][i]>tmp1['市场有效均价(主币种)'][i]:
        tmp1['优于'].iat[i]="yes"
    else:
        tmp1['优于'].iat[i]="no"
tmp1.groupby(
    pd.to_datetime(tmp1['日期']).apply(lambda x:x.month),as_index=True
    )['日期'].count()
tmp3=tmp1.groupby(
    [pd.to_datetime(tmp1['日期']).apply(lambda x:x.month),"优于"],as_index=True
    ).count().iloc[:,0:1]
tmp3.unstack()
tmp1['部门']=["haha" for i in range(len(tmp1.index))]
for i in range(len(tmp1.index)):
    if tmp1.index[i].find('新能源汽车')==-1:
        pass
    else:
        print(tmp1.index[i])
        tmp1['部门'].iat[i]="权益-私募"

len(tmp1[tmp1['部门']=="haha"].index)
set(tmp1[tmp1['部门']=="haha"].index)

#三个模式=======================
#固收
tmp1=(greatdf[
        (greatdf['业务分类']=="上交所固定收益平台") | 
        (greatdf["业务分类"]=="交易所债券借贷业务") | 
        (greatdf["业务分类"]=="交易所大宗交易") | 
        (greatdf["业务分类"]=="银行间业务") |
        (greatdf["业务分类"]=="银行间买断式回购") |
        (greatdf["业务分类"]=="银行间协议转让")]
        ).copy()
#权益
tmp1=(greatdf[
        (greatdf['业务分类']=="交易所业务") | 
        (greatdf["业务分类"]=="期权业务") | 
        (greatdf["业务分类"]=="期货业务") | 
        (greatdf["业务分类"]=="网下申购") |
        (greatdf["业务分类"]=="股转市场网上申购") |
        (greatdf["业务分类"]=="融资融券信用业务")]
        ).copy()
#场外
tmp1=(greatdf[
        (greatdf['业务分类']=="债券一级市场") | 
        (greatdf["业务分类"]=="存款类") | 
        (greatdf["业务分类"]=="开放式基金") | 
        (greatdf["业务分类"]=="开放式基金") |
        (greatdf["业务分类"]=="开放式基金") |
        (greatdf["业务分类"]=="开放式基金")]
        ).copy()
tmp1=tmp1[(tmp1['委托方向']=="正回购") | (tmp1['委托方向']=="融资回购／拆入")].copy()
tmp1=tmp1[(tmp1['委托方向']=="逆回购") | (tmp1['委托方向']=="融券回购／拆出")].copy()
tmp2=tmp1['基金名称'].apply(lambda x:True if ("鸿" in x) or ("鑫管家" in x) else False)
tmp1=tmp1[tmp2].copy()

for i in range(len(tmp1.index)):
    tmp1['指令价格(主币种)'].iat[i]=float(tmp1['指令价格(主币种)'].iat[i])
tmp1=tmp1.fillna(0)
tmp5=pd.pivot_table(tmp1,index="日期",aggfunc="count",
                        values="证券名称",columns="业务分类",margins=True,
                        margins_name='合计')
tmp5['合计'].iloc[:-1].max()
tmp2=tmp1.head()


tmp2=tmp1[['委托方向','当日成交均价(主币种)','市场有效均价(主币种)']].copy()
tmp3=pd.Series('',index=tmp2.index)
for  i in range(len(tmp2.index)):
    if tmp2.iat[i,0]=="买入" or tmp2.iat[i,0]=="卖出":
        if (tmp2.iat[i,0]=="买入" and tmp2.iat[i,1]<tmp2.iat[i,2]) or (tmp2.iat[i,0]=="卖出" and tmp2.iat[i,1]>tmp2.iat[i,2]):
            tmp3.iat[i]="yes"
        else:
            tmp3.iat[i]=r'no'
    else:
        tmp3.iat[i]='None'
len(tmp3[tmp3=='no'])
len(tmp3[tmp3=='yes'])
set(tmp1["委托方向"])
tmp4=tmp1[(tmp1["委托方向"]=="债券买入") | (tmp1["委托方向"]=="债券买入")].copy()
tmp4["当日成交金额"].sum()

#跨和正常时节的支持的
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

























