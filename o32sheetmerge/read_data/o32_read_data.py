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


def read_data_from_temp(app,dirpath):
    #dirpath=r"D:\desktop\指令全存档\temp"
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
    return greatdf




























