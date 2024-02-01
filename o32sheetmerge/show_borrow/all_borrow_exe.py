# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 14:14:55 2024

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
    tmp2=tmp1["基金名称"].map(lambda x: False if (x.find("号")!=-1)or(x.find("季季红")!=-1) else True)
    tmp1=tmp1[tmp2]
    del tmp2
    
    
    tmp1["日期"]=pd.to_datetime(tmp1["日期"])
    tmp1["月份"]=tmp1["日期"].map(lambda x:x.month)
    tmp1["指令价格(主币种)"]=tmp1["指令价格(主币种)"].astype(float)
    tmp1["到期清算金额"]=tmp1["到期清算金额"].map(lambda x:float(x.replace(",","")))
    tmp1["实际利息率"]=(tmp1["到期清算金额"]/tmp1["当日成交金额"]-1)*365
    tmp1["回购天数"]=(tmp1["实际利息率"]/tmp1["指令价格(主币种)"]*100).round(0)
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
    

    return res

def main_name(t):
    t=t.reset_index(drop=False)
    t["主机构"]=t["交易对手"].map(lambda x:x[:2])
    

    fundname=["鹏华","鹏扬",'万家','中加',"中欧","中邮","博时","嘉合",
              "嘉实","天弘","天治","富荣","格林","永赢","蜂巢","诺德",
              "鑫元","银华","银华","金鹰","长信"]
    t["主机构"]=t["主机构"].map(lambda x:x+"基金" if x in fundname else x)
    
    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)
            
    t[t["主机构"]=="自贡"]
    t["主机构"]=t["主机构"].map(lambda x:"万联证券" if x=="万联" else x)
    t["主机构"]=t["主机构"].map(lambda x:"首创证券" if x=="首创" else x)
    t["主机构"]=t["主机构"].map(lambda x:"非凡资产" if x=="非凡" else x)
    t["主机构"]=t["主机构"].map(lambda x:"长沙银行" if x=="长沙" else x)
    tmp1=t["交易对手"].map(lambda x:True if x.find("长江资管")!=-1 else False)
    t.loc[tmp1,"主机构"]="长江资管"
    t["主机构"]=t["主机构"].map(lambda x:"长江基金" if x=="长江" else x)
    t["主机构"]=t["主机构"].map(lambda x:"长安基金" if x=="长安" else x)
    t["主机构"]=t["主机构"].map(lambda x:"长城证券" if x=="长城" else x)
    t["主机构"]=t["主机构"].map(lambda x:"金信基金" if x=="金信" else x)
    t["主机构"]=t["主机构"].map(lambda x:"郑州银行" if x=="郑州" else x)
    t["主机构"]=t["主机构"].map(lambda x:"邯郸银行" if x=="邯郸" else x)
    t["主机构"]=t["主机构"].map(lambda x:"邮储银行" if x=="邮储" else x)
    t["主机构"]=t["主机构"].map(lambda x:"进出口行" if x=="进出" else x)
    t["主机构"]=t["主机构"].map(lambda x:"赣州银行" if x=="赣州" else x)
    t["主机构"]=t["主机构"].map(lambda x:"贵阳银行" if x=="贵阳" else x)
    tmp1=t["交易对手"].map(lambda x:True if x.find("西部利得")!=-1 else False)
    t.loc[tmp1,"主机构"]="西部利得基金"
    t["主机构"]=t["主机构"].map(lambda x:"西部证券" if x=="西部" else x)
    t["主机构"]=t["主机构"].map(lambda x:"蒙商银行" if x=="蒙商" else x)
    t["主机构"]=t["主机构"].map(lambda x:"萧山农商银行" if x=="萧山" else x)
    
    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)
            
    t[t["主机构"]=="江苏"]
    t["主机构"]=t["主机构"].map(lambda x:"绵阳商行" if x=="绵阳" else x)
    t["主机构"]=t["主机构"].map(lambda x:"自贡银行" if x=="自贡" else x)
    t["主机构"]=t["主机构"].map(lambda x:"福建海峡银行" if x=="福建" else x)
    t["主机构"]=t["主机构"].map(lambda x:"社保基金" if x=="社保" else x)
    t["主机构"]=t["主机构"].map(lambda x:"盛京银行" if x=="盛京" else x)
    t["主机构"]=t["主机构"].map(lambda x:"百年人寿资管" if x=="百年" else x)
    tmp1=t["交易对手"].map(lambda x:True if x.find("申万宏源")!=-1 else False)
    t.loc[tmp1,"主机构"]="申万宏源资管"
    t["主机构"]=t["主机构"].map(lambda x:"申万菱信基金" if x=="申万" else x)
    t["主机构"]=t["主机构"].map(lambda x:"瑞元资本" if x=="瑞元" else x)
    t["主机构"]=t["主机构"].map(lambda x:"灵山县农联社" if x=="灵山" else x)
    tmp1=t["交易对手"].map(lambda x:True if x.find("三湘银行")!=-1 else False)
    t.loc[tmp1,"主机构"]="湖南三湘银行"
    t["主机构"]=t["主机构"].map(lambda x:"湖南银行" if x=="湖南" else x)
    t["主机构"]=t["主机构"].map(lambda x:"渤海汇金证券" if x=="渤海" else x)
    t["主机构"]=t["主机构"].map(lambda x:"微众银行" if x=="深圳" else x)
    t["主机构"]=t["主机构"].map(lambda x:"海通证券资管" if x=="海通" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("浦发银行")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="浦发银行理财"
    tmp1=t["交易对手"].map(lambda x:True if x.find("浦发银行资管")!=-1 else False)
    t.loc[tmp1,"主机构"]="浦发银行理财"
    t["主机构"]=t["主机构"].map(lambda x:"浦发银行" if x=="浦发" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("浙商")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="浙商基金"
    t["主机构"]=t["主机构"].map(lambda x:"浙商银行" if x=="浙商" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("泰康养老")!=-1) else False)
    t.loc[tmp1,"主机构"]="浙商基金"
    t["主机构"]=t["主机构"].map(lambda x:"泰康资管" if x=="泰康" else x)
    t["主机构"]=t["主机构"].map(lambda x:"泰信基金" if x=="泰信" else x)
    t["主机构"]=t["主机构"].map(lambda x:"河北银行" if x=="河北" else x)
    t["主机构"]=t["主机构"].map(lambda x:"江西银行" if x=="江西" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("江苏银行")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="江苏银行理财"
    tmp1=t["交易对手"].map(lambda x:True if (x.find("江南农")!=-1)and(x.find("江南农")!=-1) else False)
    t.loc[tmp1,"主机构"]="江苏江南农村商业银行"
    t["主机构"]=t["主机构"].map(lambda x:"江苏银行" if x=="江苏" else x)
    
    
    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)
            
    t[t["主机构"]=="广西"]
    t["主机构"]=t["主机构"].map(lambda x:"汇添富基金" if x=="汇添" else x)
    t["主机构"]=t["主机构"].map(lambda x:"汇安基金" if x=="汇安" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("民生")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="民生银行理财"
    t["主机构"]=t["主机构"].map(lambda x:"民生银行" if x=="民生" else x)
    t["主机构"]=t["主机构"].map(lambda x:"桂林银行" if x=="桂林" else x)
    t["主机构"]=t["主机构"].map(lambda x:"柳州银行" if x=="柳州" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("杭州")!=-1)and(x.find("资管")!=-1) else False)
    t.loc[tmp1,"主机构"]="杭州银行理财"    
    t["主机构"]=t["主机构"].map(lambda x:"杭州银行" if x=="杭州" else x)
    t["主机构"]=t["主机构"].map(lambda x:"景顺长城基金" if x=="景顺" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("招商")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="招商基金"
    tmp1=t["交易对手"].map(lambda x:True if (x.find("招商")!=-1)and(x.find("养老")!=-1) else False)
    t.loc[tmp1,"主机构"]="招商基金"
    tmp1=t["交易对手"].map(lambda x:True if (x.find("招商")!=-1)and(x.find("财富")!=-1) else False)
    t.loc[tmp1,"主机构"]="招商基金"
    tmp1=t["交易对手"].map(lambda x:True if (x.find("招商")!=-1)and(x.find("证券")!=-1) else False)
    t.loc[tmp1,"主机构"]="招商证券"
    tmp1=t["交易对手"].map(lambda x:True if (x.find("招商")!=-1)and(x.find("资管")!=-1) else False)
    t.loc[tmp1,"主机构"]="招商银行资管"
    t["主机构"]=t["主机构"].map(lambda x:"招商银行" if x=="招商" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("成都")!=-1)and(x.find("农商")!=-1) else False)
    t.loc[tmp1,"主机构"]="成都农商银行"
    tmp1=t["交易对手"].map(lambda x:True if (x.find("成都")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="成都银行理财"
    t["主机构"]=t["主机构"].map(lambda x:"成都银行" if x=="成都" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("恒丰")!=-1)and(x.find("资管")!=-1) else False)
    t.loc[tmp1,"主机构"]="恒丰银行理财"    
    t["主机构"]=t["主机构"].map(lambda x:"恒丰银行" if x=="恒丰" else x)
    t["主机构"]=t["主机构"].map(lambda x:"徽商银行理财" if x=="徽商" else x)
    t["主机构"]=t["主机构"].map(lambda x:"德邦基金" if x=="德邦" else x)
    t["主机构"]=t["主机构"].map(lambda x:"弘业期货" if x=="弘业" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("建信")!=-1)and(x.find("信托")!=-1) else False)
    t.loc[tmp1,"主机构"]="建信信托"    
    tmp1=t["交易对手"].map(lambda x:True if (x.find("建信")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="建信基金"    
    tmp1=t["交易对手"].map(lambda x:True if (x.find("建信")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="建信资本"    
    t["主机构"]=t["主机构"].map(lambda x:"广银理财" if x=="广银" else x)
    
    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)
            
    t[t["主机构"]=="富国"]
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广西")!=-1)and(x.find("银行")!=-1) else False)
    t.loc[tmp1,"主机构"]=t.loc[tmp1,"交易对手"]  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广州农")!=-1)and(x.find("银行")!=-1) else False)
    t.loc[tmp1,"主机构"]="广州农村商业银行"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广州")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="广州银行理财"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广发")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="广发基金"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广发")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="广发银行理财"  
    t["主机构"]=t["主机构"].map(lambda x:"广州银行" if x=="广州" else x)
    t["主机构"]=t["主机构"].map(lambda x:"广发银行" if x=="广发" else x)
    t["主机构"]=t["主机构"].map(lambda x:"汇添富基金" if x=="广东" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广东")!=-1)and(x.find("南海")!=-1) else False)
    t.loc[tmp1,"主机构"]="广东南海农商行"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("广东")!=-1)and(x.find("顺德")!=-1) else False)
    t.loc[tmp1,"主机构"]="广东顺德农商行"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安")!=-1)and(x.find("信托")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安信托"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安基金"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安银行理财"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安")!=-1)and(x.find("财富")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安基金"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安")!=-1)and(x.find("资产")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安资产"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安")!=-1)and(x.find("财险")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安财险"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("平安证券")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安证券资管"   
    t["主机构"]=t["主机构"].map(lambda x:"平安银行" if x=="平安" else x)
    t["主机构"]=t["主机构"].map(lambda x:"工银瑞信基金" if x=="工银" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("工行")!=-1)and(x.find("富国")!=-1) else False)
    t.loc[tmp1,"主机构"]="富国基金"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("工行")!=-1)and(x.find("工银瑞信")!=-1) else False)
    t.loc[tmp1,"主机构"]="工银瑞信基金"  
    t["主机构"]=t["主机构"].map(lambda x:"山西证券资管" if x=="山证" else x)
    t["主机构"]=t["主机构"].map(lambda x:"山西证券资管" if x=="山西" else x)
    t["主机构"]=t["主机构"].map(lambda x:"江苏江南农村商业银行" if x=="富江" else x)
    
    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)
            
    t[t["主机构"]=="国投"]   
    t["主机构"]=t["主机构"].map(lambda x:"富国基金" if x=="富国" else x)
    t["主机构"]=t["主机构"].map(lambda x:"国海富兰克林基金" if x=="富兰" else x)
    t["主机构"]=t["主机构"].map(lambda x:"容县农联社" if x=="容县" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("安信")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="安信基金"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("安信")!=-1)and(x.find("资管")!=-1) else False)
    t.loc[tmp1,"主机构"]="安信资管"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("宁波")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="宁波银行理财"   
    t["主机构"]=t["主机构"].map(lambda x:"宁波银行" if x=="宁波" else x)
    t["主机构"]=t["主机构"].map(lambda x:"天风证券" if x=="天风" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("大连")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="大连农商银行理财"       
    t["主机构"]=t["主机构"].map(lambda x:"大连银行" if x=="大连" else x)    
    t["主机构"]=t["主机构"].map(lambda x:"基本养老保险基金" if x=="基本" else x)
    t["主机构"]=t["主机构"].map(lambda x:"圆信永丰基金" if x=="圆信" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联安鑫")!=-1)and(x.find("国联安鑫")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联证券资管"       
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联安昕")!=-1)and(x.find("国联安")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联证券资管"       
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联安")!=-1)and(x.find("国联安")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联安基金"       
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联")!=-1)and(x.find("指数基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联基金"       
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联证券资管"       
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联")!=-1)and(x.find("集合")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联证券资管"       
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国联")!=-1)and(x.find("金如意")!=-1) else False)
    t.loc[tmp1,"主机构"]="国联证券资管"    
    t["主机构"]=t["主机构"].map(lambda x:"国联证券" if x=="国联" else x)
    t["主机构"]=t["主机构"].map(lambda x:"国海证券资管" if x=="国海" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("国泰君安")!=-1)and(x.find("证券")==-1) else False)
    t.loc[tmp1,"主机构"]="国泰君安证券资管"   
    t["主机构"]=t["主机构"].map(lambda x:"国泰君安证券" if x=="国泰" else x)
    
    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)
            
    t[t["主机构"]=="华泰"]   
    t["主机构"]=t["主机构"].map(lambda x:"国投瑞银基金" if x=="国投" else x)
    t["主机构"]=t["主机构"].map(lambda x:"国泰君安证券资管" if x=="国君" else x)
    t["主机构"]=t["主机构"].map(lambda x:"四川长宁竹海农商行" if x=="四川" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("吉林省农联社")!=-1)and(x.find("吉林省农联社")!=-1) else False)
    t.loc[tmp1,"主机构"]="吉林省农联社"   
    t["主机构"]=t["主机构"].map(lambda x:"吉林银行" if x=="吉林" else x)
    t["主机构"]=t["主机构"].map(lambda x:"合浦县农联社" if x=="合浦" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("厦门农商")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="厦门农商银行理财"   
    t["主机构"]=t["主机构"].map(lambda x:"厦门国际银行" if x=="厦门" else x)
    t["主机构"]=t["主机构"].map(lambda x:"博白县农联社" if x=="博白" else x)
    t["主机构"]=t["主机构"].map(lambda x:"南京银行" if x=="南京" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华西证券资管" if x=="华西" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华润信托" if x=="华润" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("华泰柏瑞")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="华泰柏瑞基金"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("华泰资产")!=-1)and(x.find("资产")!=-1) else False)
    t.loc[tmp1,"主机构"]="华泰资产"  
    t["主机构"]=t["主机构"].map(lambda x:"华泰证券" if x=="华泰" else x)

    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)       
            
    t[t["主机构"]=="交银"]   
    t["主机构"]=t["主机构"].map(lambda x:"华富基金" if x=="华富" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华宝兴业货币基金" if x=="华宝" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华安基金" if x=="华安" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华夏银行" if x=="华夏" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华商基金" if x=="华商" else x)
    t["主机构"]=t["主机构"].map(lambda x:"华创证券" if x=="华创" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("北京")!=-1)and(x.find("资管")!=-1) else False)
    t.loc[tmp1,"主机构"]="北京银行资管"   
    t["主机构"]=t["主机构"].map(lambda x:"北京银行" if x=="北京" else x)
    t["主机构"]=t["主机构"].map(lambda x:"创金合信基金" if x=="创金" else x)
    t["主机构"]=t["主机构"].map(lambda x:"泰康资产" if x=="农行" else x)
    t["主机构"]=t["主机构"].map(lambda x:"农业银行" if x=="农业" else x)
    t["主机构"]=t["主机构"].map(lambda x:"内蒙古银行" if x=="内蒙" else x)
    t["主机构"]=t["主机构"].map(lambda x:"兴银理财" if x=="兴银" else x)
    t["主机构"]=t["主机构"].map(lambda x:"兴全基金" if x=="兴证" else x)
    t["主机构"]=t["主机构"].map(lambda x:"兴全基金" if x=="兴全" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("兴业")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="兴业基金"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("兴业")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="兴银理财"  
    t["主机构"]=t["主机构"].map(lambda x:"兴业银行" if x=="兴业" else x)
    t["主机构"]=t["主机构"].map(lambda x:"全国社保基金" if x=="全国" else x)
    t["主机构"]=t["主机构"].map(lambda x:"光大证券资管" if x=="光证" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("光大永明")!=-1)and(x.find("光大永明")!=-1) else False)
    t.loc[tmp1,"主机构"]="光大永明资产"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("光大")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="光大银行理财"  
    tmp1=t["交易对手"].map(lambda x:True if (x.find("光大")!=-1)and(x.find("资管")!=-1) else False)
    t.loc[tmp1,"主机构"]="光大证券资管"      
    t["主机构"]=t["主机构"].map(lambda x:"光大银行" if x=="光大" else x)
    t["主机构"]=t["主机构"].map(lambda x:"信银理财" if x=="信银" else x)
    t["主机构"]=t["主机构"].map(lambda x:"信达基金" if x=="信达" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("交银施罗德")!=-1)and(x.find("交银施罗德")!=-1) else False)
    t.loc[tmp1,"主机构"]="交银施罗德基金"      
    t["主机构"]=t["主机构"].map(lambda x:"人保资产" if x=="人保" else x)
    t["主机构"]=t["主机构"].map(lambda x:"交银理财" if x=="交银" else x)

    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)       
            
    t[t["主机构"]=="东吴"]   
    t["主机构"]=t["主机构"].map(lambda x:"交通银行" if x=="交通" else x)
    t["主机构"]=t["主机构"].map(lambda x:"平安资产" if x=="交行" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("五矿信托")!=-1)and(x.find("五矿信托")!=-1) else False)
    t.loc[tmp1,"主机构"]="五矿信托"     
    t["主机构"]=t["主机构"].map(lambda x:"五矿证券" if x=="五矿" else x)
    t["主机构"]=t["主机构"].map(lambda x:"云南红塔银行" if x=="云南" else x)
    t["主机构"]=t["主机构"].map(lambda x:"乌鲁木齐银行" if x=="乌鲁" else x)
    t["主机构"]=t["主机构"].map(lambda x:"中银富登村镇银行" if x=="中银" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中金")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="中金基金"     
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中金")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="中金基金"     
    t["主机构"]=t["主机构"].map(lambda x:"中金公司" if x=="中金" else x)
    t["主机构"]=t["主机构"].map(lambda x:"中信证券资管" if x=="中行" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中航工业集团财务")!=-1)and(x.find("中航工业集团财务")!=-1) else False)
    t.loc[tmp1,"主机构"]="中航工业集团财务"     
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中航")!=-1)and(x.find("信托")!=-1) else False)
    t.loc[tmp1,"主机构"]="中航信托"   
    t["主机构"]=t["主机构"].map(lambda x:"中融基金" if x=="中融" else x)
    t["主机构"]=t["主机构"].map(lambda x:"中航基金" if x=="中航" else x)
    t["主机构"]=t["主机构"].map(lambda x:"中金公司" if x=="中金" else x)
    t["主机构"]=t["主机构"].map(lambda x:"中泰证券资管" if x=="中泰" else x)
    t["主机构"]=t["主机构"].map(lambda x:"平安资产" if x=="中央" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中国人保")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="人保资产"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中国")!=-1)and(x.find("平安")!=-1) else False)
    t.loc[tmp1,"主机构"]="平安资产"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中国")!=-1)and(x.find("长城")!=-1) else False)
    t.loc[tmp1,"主机构"]="中国长城资管"   
    t["主机构"]=t["主机构"].map(lambda x:"中国银行" if x=="中国" else x)
    t["主机构"]=t["主机构"].map(lambda x:"中原银行" if x=="中原" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信")!=-1)and(x.find("信托")!=-1) else False)
    t.loc[tmp1,"主机构"]="中信信托"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信证券")!=-1)and(x.find("号")!=-1) else False)
    t.loc[tmp1,"主机构"]="中信证券资管"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信证券")!=-1)and(x.find("计划")!=-1) else False)
    t.loc[tmp1,"主机构"]="中信证券资管"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信证券")!=-1)and(x.find("养老金")!=-1) else False)
    t.loc[tmp1,"主机构"]="中信证券资管"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信保诚")!=-1)and(x.find("保险")!=-1) else False)
    t.loc[tmp1,"主机构"]="中信保诚人寿"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信保诚")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="中信保诚基金"   
    tmp1=t["交易对手"].map(lambda x:True if (x.find("中信")!=-1)and(x.find("理财")!=-1) else False)
    t.loc[tmp1,"主机构"]="信银理财" 
    t["主机构"]=t["主机构"].map(lambda x:"中信证券" if x=="中信" else x)
    t["主机构"]=t["主机构"].map(lambda x:"东证融汇" if x=="东证" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("东莞")!=-1)and(x.find("农村")!=-1) else False)
    t.loc[tmp1,"主机构"]="东莞农村商业银行" 
    t["主机构"]=t["主机构"].map(lambda x:"东莞银行" if x=="东莞" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("东海")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="东海基金" 
    t["主机构"]=t["主机构"].map(lambda x:"东海证券" if x=="东海" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("东方红")!=-1)and(x.find("东方红")!=-1) else False)
    t.loc[tmp1,"主机构"]="东证资管" 
    tmp1=t["交易对手"].map(lambda x:True if (x.find("东方")!=-1)and(x.find("基金")!=-1) else False)
    t.loc[tmp1,"主机构"]="东方基金" 
    
    t["主机构"]=t["主机构"].map(lambda x:"东方证券" if x=="东方" else x)

    for i in t["主机构"].unique():
        if len(i)==2:
            print(i)       
            
    t[t["主机构"]=="上农"]   
    t["主机构"]=t["主机构"].map(lambda x:"东吴人寿" if x=="东吴" else x)
    t["主机构"]=t["主机构"].map(lambda x:"东兴证券" if x=="东兴" else x)
    t["主机构"]=t["主机构"].map(lambda x:"东亚前海证券" if x=="东亚" else x)
    t["主机构"]=t["主机构"].map(lambda x:"上银基金" if x=="上银" else x)
    t["主机构"]=t["主机构"].map(lambda x:"东亚前海证券" if x=="东亚" else x)
    tmp1=t["交易对手"].map(lambda x:True if (x.find("上海")!=-1)and(x.find("期")!=-1) else False)
    t.loc[tmp1,"主机构"]="上海农商银行理财" 
    t["主机构"]=t["主机构"].map(lambda x:"上海农商银行" if x=="上海" else x)
    t["主机构"]=t["主机构"].map(lambda x:"上海农商银行理财" if x=="上农" else x)
    
    return t
    

















