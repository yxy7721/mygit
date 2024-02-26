# -*- coding: utf-8 -*-
"""
Created on Thu Feb 16 10:22:24 2023

@author: yangxy
"""

import pandas as pd
import numpy as np
import copy
import openpyxl as op
import os
import xlwings as xw

app=xw.App(visible=False,add_book=False)
app.display_alerts=False
app.screen_updating=False
#以上的初始必做

def read_one_after_one(dirpath):
    dirlist=os.listdir(dirpath)
    greatlis=dict()
    filenamelis=list()
    for filename in dirlist:
        print(filename)
        filenamelis.append(filename)
        if os.path.isdir(os.path.join(dirpath,filename)) or filename=="ok.xlsx":
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
            df.index=pd.to_datetime(df['date'])
            del df['date']
            #df=df.resample('3M',axis=0,closed="right",label="right").last()
            df=df.reset_index(drop=False)
            greatdf[wb.sheets[she].name]=df
        greatlis[filename]=greatdf
        wb.close()
    return greatlis,filenamelis

def remove_same(greatlis):
    tmp1=set()
    for i in greatlis.keys():
        tmp1=set(greatlis[i].keys()) | tmp1
    shtlis=copy.deepcopy(tmp1)
    del tmp1
    mydict=dict()
    for sht in shtlis:
        sht
        df=pd.DataFrame(columns=["date"])
        for gtky in greatlis.keys():
            sce=greatlis[gtky]
            #break
            #greatlis.keys()
            #sce=greatlis[]
            print(sce.keys())
            if sht in sce.keys():
                tmp1=copy.deepcopy(sce[sht])
                tmp1.index=pd.to_datetime(tmp1['date'])
                #tmp1=tmp1.resample('3M',axis=0,closed="right",label="right").last()
                tmp1['date']=tmp1.index
                tmp1.index.name='index'
                df=pd.merge(df,tmp1,how="outer",on="date")
                tmp1=df.copy()
                i=sht
                tmp2=pd.Series(tmp1.columns).apply(lambda x:True if x[-2:-1]=="_" else False)
                if len(tmp2[tmp2])>0:
                    tmp3=tmp1.loc[:,tmp1.columns[tmp2]]
                    tmp4=pd.Series(tmp3.columns).apply(lambda x:x[:-2])
                    tmp4=set(tmp4)
                    tmp7=pd.DataFrame(index=tmp1.index)
                    for j in tmp4:
                        #break
                        tmp5=tmp3.loc[:,j+"_x"]
                        tmp6=tmp3.loc[:,j+"_y"]
                        tmp5=tmp5.replace(0,np.nan)
                        tmp6=tmp6.replace(0,np.nan)
                        for k in range(len(tmp5)):
                            #break
                            if (tmp5[k]!=tmp5[k]) or (tmp5[k]==""):
                                tmp5[k]=tmp6[k]
                        tmp5=pd.DataFrame(tmp5,index=tmp1.index)
                        tmp5.columns=[j]
                        tmp7=pd.concat([tmp7,tmp5],axis=1)
                    tmp1=tmp1.loc[:,tmp1.columns[tmp2.apply(lambda x:not(x))]]
                    del tmp2,tmp3,tmp4,tmp5,tmp6
                    tmp7=pd.concat([tmp1['date'],tmp7],axis=1)
                    tmp1=pd.concat([tmp7,tmp1.iloc[:,1:]],axis=1)
                    df=tmp1.copy()
            else:
                pass
        df=df.sort_values(by='date',axis=0,ascending=True,inplace=False,na_position='last')
        df=df.reset_index(drop=True)
        mydict[sht]=df
    del sht,sce,tmp1
    return mydict
    
def show_side_number(greatlis,axis="both"):
    greatlis.keys()
    for filename in greatlis.keys():
        filename
        for shtname in greatlis[filename].keys():
            shtname
            greatlis[filename][shtname]
            print(filename,shtname,len(greatlis[filename][shtname].columns),
                  len(greatlis[filename][shtname].index))

def just_pile(greatlis,ept=["hkadj.xlsx"]):
    sheetset=set()
    pile_up_dict=dict()
    ept_dict=dict()
    for filename in greatlis.keys():
        sheetset=sheetset | set(greatlis[filename].keys())
    for filename in greatlis.keys():
        print(filename)
        if filename in ept:
            ept_dict[filename]=greatlis[filename].copy()
            continue
        #filename='stockclose2018.xlsx' 
        for sheetname in sheetset:
            if sheetname in pile_up_dict.keys():
                tmp1=pile_up_dict[sheetname].copy()
                tmp2=greatlis[filename][sheetname].copy()
                pile_up_dict[sheetname]=pd.concat([tmp1,tmp2],axis=0)
            else:
                pile_up_dict[sheetname]=greatlis[filename][sheetname].copy()
    for i in pile_up_dict.keys():
        pile_up_dict[i].reset_index(drop=True)
    pile_up_dict={"smooth":pile_up_dict}
    return pile_up_dict,ept_dict

def middle_zero(eptlis):
    eptlis
    for i in eptlis.keys():
        tmp1=eptlis[i]
        for j in tmp1.keys():
            tmp2=tmp1[j]
            for col in tmp2.columns:
                tmp3=tmp2[col]
                f=0
                while True:
                    if tmp3.iat[f]==0:
                        f=f+1
                        if f==len(tmp3):
                            break
                    else:
                        break
                if f==len(tmp3):
                    continue
                while True:
                    if tmp3.iat[f]!=0:
                        f=f+1
                        if f==len(tmp3):
                            break
                    else:
                        break
                if f!=len(tmp3):
                    print("it has middle zero!!!",col)
                
    del tmp1,tmp2,tmp3,f,i,j,col
                
    
def replace_middle_zero(eptlis):
    for fname in eptlis.keys():
        for shtname in eptlis[fname]:
            the_sht=eptlis[fname][shtname]
            for col in the_sht.columns:
                tmp1=the_sht[col]
                f=0
                while True:
                    if tmp1.iat[f]==0:
                        f=f+1
                        if f==len(tmp1):
                            break
                    else:
                        break
                if f==len(tmp1):
                    continue
                while True:
                    if tmp1.iat[f]!=0:
                        f=f+1
                        if f==len(tmp1):
                            break
                    else:
                        break
                while True:
                    if f==len(tmp1):
                        break
                    if tmp1.iat[f]==0:
                        print("here",f)
                        nextf=f
                        while True:
                            if tmp1.iat[nextf]==0:
                                nextf=nextf+1
                                if nextf==len(tmp1):
                                    break
                            else:
                                break
                        if nextf!=len(tmp1):
                            tmp1.iat[f]=(tmp1.iat[f-1]+tmp1.iat[nextf])/2
                            f=f+1
                        elif tmp1.iat[nextf]==0.0:
                            tmp1.iat[f]=tmp1.iat[f-1]
                            f=f+1
                    else:
                        f=f+1
    return eptlis
                
def slice_into_pile(mydict,how_many_pile=6):
    how_many_pile=6
    piles=list()
    for tmp1,tmp2 in mydict.items():
        for tmp3,tmp4 in tmp2.items():
            type(tmp4)
            slice_number_list=np.around(np.linspace(0,len(tmp4.index),how_many_pile),decimals=0)
            for i in range(1,len(slice_number_list)):
                slice_point=int(slice_number_list[i])
                piles.append(tmp4.iloc[int(slice_number_list[i-1]):slice_point,:])
    return piles
       
def save_pile_into_files(mydict):
    f=0
    for i in mydict:
        shuchu=app.books.add()
        shuchu.sheets[0].name
        shuchu.sheets[0].used_range.last_cell.row
        shuchu.sheets[0].used_range.last_cell.column
        shuchu.sheets[0].range("A1").options(pd.DataFrame, index=False).value=i
        f=f+1
        shuchu.save(r'D:\desktop\checkNwash\close_washed{}.xlsx'.format(str(f)))
        '''
        shuchu.sheets[0].range(
            (1,2),(len(tmp2.index)+1,len(tmp2.columns))
            ).options(pd.DataFrame, index=False).api.NumberFormat = "@"
        '''
        
        
    
         

#读取并去中间的零
dirpath=r"D:\desktop\mydatabase\adjfac"
greatlis,filenamelis=read_one_after_one(dirpath)
show_side_number(greatlis)
greatlis,eptlis=just_pile(greatlis,ept=["null"])
show_side_number(greatlis)
show_side_number(eptlis)
middle_zero(eptlis)
middle_zero(greatlis)


for i in eptlis.keys():
    greatlis[i]=eptlis[i]
del eptlis,i
mydict=remove_same(greatlis)
del greatlis,dirpath,filenamelis
show_side_number({"smooth":mydict})
middle_zero({"smooth":mydict})
mydict=replace_middle_zero({"smooth":mydict})
middle_zero(mydict)

#横向切割后保存
mydict=slice_into_pile(mydict,how_many_pile=6)
save_pile_into_files(mydict)


'''
shuchu=app.books.add()
for sh in range(10):
    print(sh)
    dataget=pd.DataFrame(columns=fundpool.iloc[:,0],index=kk)
    for i in range(len(dataget.index)):
        for j in range(len(dataget.columns)):
            dataget.iat[i,j]=str(i+2)+","+str(j+2)
    dataget=dataget.applymap(lambda x: '=f_prt_heavilyheldstocktonav(indirect(address(1,{},4)),indirect(address({},1,4)),{})'.format(x.split(',')[1],x.split(',')[0],sh+1))
    dataget.index.name='date'
    shuchu.sheets.add()
    shuchu.sheets[0].name="第"+str(sh+1)+"大重仓股占比"
    shuchu.sheets[0].range('A1').options(pd.DataFrame, index=True).value=dataget
shuchu.save(r'D:\desktop\winddataget\重仓占比取数.xlsx')
'''

#四、列出不全的标的，保存数据完整的标的列表。
def iscomplete(se):
    #先剔除前面的空值，因为基金还没成立，这很正常
    pin="reset"
    for i in range(len(se)):
        if se.iat[i]==se.iat[i] and not(se.iat[i] is None):
            pin=i
            break
        else:
            "它是空值"
    if (i==len(se)-1) and (pin!=i):
        return "null"
    for i in range(pin,len(se)):
        if se.iat[i]==se.iat[i] and not(se.iat[i] is None):
            "它是实值"
        else:
            pin="breaked"
            break
    return pin
    
okayfundlis=dict()
badfundlis=dict()
for shtname,df in mydict.items():
    #break
    #mydict.keys()
    #df=mydict[shtname]
    oklis=list()
    badlis=list()
    for col in df.columns:
        if col=="date":
            continue
        if iscomplete(df[col])=="null":
            print("null",shtname,col)
            #badlis.append(col)
        elif iscomplete(df[col])=="breaked":
            print("breaked",shtname,col)
            badlis.append(col)
        else:
            oklis.append(col)
    okayfundlis[shtname]=oklis
    badfundlis[shtname]=badlis
del oklis,badlis
greatset=set()
for i in okayfundlis.values():
    greatset=set(i) | greatset
miniset=greatset
for i in okayfundlis.values():
    miniset=set(i) & miniset
okayfundlis=pd.DataFrame(miniset)
okayfundlis.columns=["fundname"]
#把ok的输出
shuchu=app.books.add()
shuchu.sheets.add()
shuchu.sheets[0].name="ok"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=okayfundlis
shuchu.save(r'D:\desktop\checkNwash\ok.xlsx')
del col,df,greatset,i,miniset,mydict,okayfundlis,shtname
#把badguy输出
greatset=set()
for i in badfundlis.values():
    greatset=set(i) | greatset
badguy=pd.DataFrame(greatset)
badguy.columns=["fundname"]
shuchu=app.books.add()
shuchu.sheets.add()
shuchu.sheets[0].name="badguy"
shuchu.sheets[0].range('A1').options(pd.DataFrame, index=False).value=badguy
shuchu.save(r'D:\desktop\checkNwash\badguy.xlsx')
        
app.kill()
































