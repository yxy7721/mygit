# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
import sys
import pyecharts
from selenium import webdriver
import copy
#import webbrowser
import os

def jiaoyan(haha):
    if len(set(haha['code']))<len(haha['code']):
        return False
    else:
        return True

def pdlevel():
    
    lvlscale=[]
    lvlscale.append([df[df['parent'] == 0]['code'][0]])
    level=2
    
    while True:
        nextnodelis=[]
        for i in range(len(lvlscale[level-2])):
            nextnodelis=nextnodelis+list(df[df['parent'] == lvlscale[level-2][i]]['code'])
        if nextnodelis==[]:
            break
        lvlscale.append(nextnodelis)
        level=level+1
    level=level-1
    return lvlscale
            
def pdchild():
    whosmychild=[-1 for i in range(len(df.index))]
    for i in range(len(df.index)):
        code=df['code'][i]
        flag=list(df[df['parent']==code]['code'])
        whosmychild[i]=flag
    return whosmychild

def makedict(dic,code):
    print(code)

    if len(df[df['parent']==code].index)==0:
        pass
    else:
        dic['children']=[]
        for next_code in list(df[df['parent']==code]['code']):
            newdict=dict()
            newdict['name']=df[df['code']==next_code]['value'].values[0]
            dic['children'].append((makedict(newdict,next_code)))
    
    return dic

def makedf(dic,parent):
    global code
    childlist=dic.get('children')
    new=pd.DataFrame(None,index=[0],columns=['code','value','parent'])
    new['value'][0]=dic['name']
    new['parent']=parent
    new['code']=code
    if childlist==None:
        return new
    childcode=code
    for i in range(len(childlist)):
        code=code+1
        nextnew=makedf(childlist[i],childcode)
        new=pd.concat([new,nextnew],axis=0)
    return new

def partdata(dic,nowlevel):

    if nowlevel>=N:
        if dic.get('children') is None:
            pass
        else:
            dic.pop('children')
    else:
        if dic.get('children') is None:
            pass
        else:
            for i in range(len(dic['children'])):
                partdata(dic['children'][i],nowlevel+1)
   
def marknumber1(diclist):
    for i in range(len(diclist)):
        diclist[i]['mark']=str(i+1)
        if diclist[i].get('children')==None:
            continue
        marknumber1(diclist[i]['children'])
    
def marknumber2(dic):
    if dic.get('children')==None:
        return
    for i in range(len(dic['children'])):
        dic['children'][i]['mark']=dic['mark']+"."+dic['children'][i]['mark']
        if dic.get('children')==None:
            continue
        marknumber2(dic['children'][i])
        
def markpaste(diclist):
    for i in range(len(diclist)):
        diclist[i]['name']=diclist[i]['mark']+diclist[i]['name']
        if  diclist[i].get('children')==None:
            continue
        markpaste(diclist[i]['children'])
    
def searchmark(diclist):    
    for i in range(len(marklist)-1):
        diclist=diclist[int(marklist[i])-1]['children']
    
    return diclist[int(marklist[-1])-1]
        
def broins(diclist):

    for i in range(len(insposlist)):
        if i==len(insposlist)-1:
            break
        diclist=diclist[int(insposlist[i])-1]['children']
    flag=0
    while True:
        if int(insposlist[-1])<int(diclist[flag]['mark'].split('.')[-1]):
            break
        flag=flag+1
        if flag>len(diclist)-1:
            break
    charu=dict()
    charu['name']=insname
    diclist.insert(flag,charu)    

def dadins(diclist):
    
    for i in range(len(fatherposlist)):
        if i==len(fatherposlist)-1:
            break
        diclist=diclist[int(fatherposlist[i])-1]['children']
    charu=dict()
    charu['name']=insname
    diclist[int(fatherposlist[-1])-1]['children']=[charu]

def swiftnode(diclist):
    conductlist=diclist
    for i in range(len(startposlist)):
        if i==len(startposlist)-1:
            break
        if i==len(startposlist)-2:
            cidiceng=conductlist[int(startposlist[i])-1]
        conductlist=conductlist[int(startposlist[i])-1]['children']
    yidong=copy.deepcopy(conductlist[int(startposlist[-1])-1])
    if len(conductlist)==1:
        del cidiceng['children']
    else:
        del diclist[int(startposlist[-1])-1]
    
    conductlist=diclist
    for i in range(len(endpos_father_list)):
        if i==len(endpos_father_list)-1:
            break
        conductlist=conductlist[int(endpos_father_list[i])-1]['children']
    if conductlist[int(endpos_father_list[-1])-1].get('children') is None:
        conductlist[int(endpos_father_list[-1])-1]['children']=[yidong]
    else:
        conductlist[int(endpos_father_list[-1])-1]['children'].append([yidong])
        
def delnode(diclist):
    for i in range(len(delposlist)):
        if i==len(delposlist)-1:
            break
        if i==len(delposlist)-2:
            cidiceng=diclist[int(delposlist[i])-1]
        diclist=diclist[int(delposlist[i])-1]['children']
    if diclist[int(delposlist[-1])-1].get('children') is None:
        if len(cidiceng['children'])>1:
            del cidiceng['children'][int(delposlist[-1])-1]
        else:
            del cidiceng['children']
    else:
        del diclist[int(delposlist[-1])-1]
        
def modify(diclist):
    for i in range(len(modifylist)):
        if i==len(modifylist)-1:
            break
        diclist=diclist[int(modifylist[i])-1]['children']
    diclist[int(modifylist[-1])-1]['name']=modifyname

def showoff(data):
    newcharts = (
        pyecharts.charts.Tree()
        .add("", data,initial_tree_depth=-1)
        .render(r"D:\desktop\ymind\new.html")
    )
    
    
    
    #driver.get(newcharts) 

#A1.读取原文件(初始化必做)
df=pd.read_csv(r'D:\desktop\ymind\xmind.csv', encoding='gbk')
#if jiaoyan(df) is False:
    #sys.exit(0)

#A2.原文件整理成字典（初始化必做）
data=[dict()]
data[0]['name']=df[df['parent']==0]['value'].values[0]
data=[makedict(data[0],1)]
marknumber1(data)
marknumber2(data[0])

#B.字典再转换成文件进行保存（若有修改则必做）
code=1
newdf=makedf(data[0],0)
newdf.index=range(len(newdf.index))
newdf.to_csv(r'D:\desktop\ymind\xmind.csv',index=False,encoding='gbk')



#os.system(r'D:\desktop\ymind\\' "new.html")
#webbrowser.open(r'D:\desktop\ymind\new.html')
#print(webbrowser.get())

#C1.进行部分节点数据截取以供展示
#仅展示前N层节点
N=4
dic=copy.deepcopy(data[0])
partdata(dic,1)

showdata=[copy.deepcopy(dic)]
markpaste(showdata)
#driver = webdriver.Chrome(r"D:\chromedriver\chromedriver.exe")
showoff(showdata)
os.system(r'D:\desktop\ymind\\' "new.html")

#del driver

#输入节点字符串，直接摘出来节点及之后节点
markinput="1"
marklist=markinput.split('.')
dic=copy.deepcopy(data[0])
dic=searchmark([dic])

showdata=[copy.deepcopy(dic)]
markpaste(showdata)

#driver = webdriver.Chrome(r"D:\chromedriver\chromedriver.exe")
showoff(showdata)
os.system(r'D:\desktop\ymind\\' "new.html")
#del driver



#C2.删除节点
delpos='1.1.2.1.1'
delposlist=delpos.split('.')
diclist=copy.deepcopy(data)
delnode(diclist)
data=copy.deepcopy(diclist)
marknumber1(data)   
marknumber2(data[0])   #重新加mark

#C3.修改节点文字内容
modifypos='1.1.1.5.7.4'
modifyname='修改悬挂' #left_indent、right_indent、first_line_indent
modifylist=modifypos.split('.')
diclist=copy.deepcopy(data)
modify(diclist)
data=copy.deepcopy(diclist)
marknumber1(data)   
marknumber2(data[0])   #重新加mark

#C4.部分节点父子关系重构
startpos='1.1.1.5.7.3.1'
endpos_father='1.1.1.5.7.3'
startposlist=startpos.split('.')
endpos_father_list=endpos_father.split('.')
diclist=copy.deepcopy(data)
swiftnode(diclist)
data=copy.deepcopy(diclist)
marknumber1(data)   
marknumber2(data[0])   #重新加mark

#C5.添加兄弟节点
inspos='1.1.2.4'
insname='自动生成wind取数工具人'
insposlist=inspos.split('.')
diclist=copy.deepcopy(data)
broins(diclist)
data=copy.deepcopy(diclist)
marknumber1(data)   
marknumber2(data[0])   #重新加mark

#C5.添加子节点（必须无兄弟才能使用）
fatherpos2ins='1.1.3'
insname='国君宏观体系学习'
fatherposlist=fatherpos2ins.split('.')
diclist=copy.deepcopy(data)
dadins(diclist)
data=copy.deepcopy(diclist)
marknumber1(data)   
marknumber2(data[0])   #重新加mark

'''
data = [
    {
        "children": [
            {"name": "B"},
            {
                "children": [{"children": [{"name": "I"}], "name": "E"}, {"name": "F"}],
                "name": "C",
            },
            {
                "children": [
                    {"children": [{"name": "J"}, {"name": "K"}], "name": "G"},
                    {"name": "H"},
                ],
                "name": "D",
            },
        ],
        "name": "A",
    }
]
'''


#D.生成交互html
newcharts = (
    pyecharts.charts.Tree()
    .add("", data,initial_tree_depth=-1)
    .render(r"D:\desktop\ymind\new.html")
)


driver = webdriver.Chrome(r"D:\chromedriver\chromedriver.exe")
driver.get(newcharts) 




















