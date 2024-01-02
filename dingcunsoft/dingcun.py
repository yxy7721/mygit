# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from PyQt5.QtWidgets import QApplication,QLabel,QWidget,QTextEdit,QGridLayout,QPushButton,QLineEdit,QTableView,QHeaderView
from PyQt5.QtCore import QRect,Qt
from PyQt5.QtGui import QStandardItemModel,QStandardItem,QFont
#from PyQt5.QtCore import Qt
import xlwings as xw
import pandas as pd
import sys
from time import strftime
from datetime import datetime
from os import system

class TextEditDemo(QWidget):
    def __init__(self, parent=None):
        super(TextEditDemo, self).__init__(parent)
        self.initUI()
    
    def initUI(self):   
        self.setWindowTitle('定存')

        #定义窗口的初始大小
        self.resize(1000,1000)
                
        self.centralwidget = QWidget()
        self.centralwidget.setObjectName("centralwidget")
        
        self.pushButton = QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QRect(150, 640, 131, 40))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.setText("识别")

        self.pushButton_2 = QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QRect(430, 640, 131, 40))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setText("清空")
        
        self.pushButton_3 = QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QRect(800, 640, 131, 40))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.setText("纳入")
                
        self.pushButton_4 = QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QRect(150, 900, 131, 40))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.setText("导入")
        
        self.pushButton_5 = QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QRect(430, 900, 131, 40))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_5.setText("导出")
        
        self.pushButton_6 = QPushButton(self.centralwidget)
        self.pushButton_6.setGeometry(QRect(720, 900, 131, 40))
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_6.setText("清色导出")

        self.labz = QLabel(self.centralwidget)
        self.labz.setText('这里是消息框')
        self.labz.setFont(QFont("Times" , 16))
        self.labz.setGeometry(QRect(310, 160, 400, 360))
        self.labz.setWordWrap(True)
        self.labz.setAlignment(Qt.AlignTop)

        self.lineday = QLineEdit(self.centralwidget)
        self.lineday.clear()
        self.lineday.setGeometry(QRect(310, 540, 180, 40))
        self.lineday.setText(datetime.now().strftime("%Y/%m/%d"))

        self.linever = QLineEdit(self.centralwidget)
        self.linever.clear()
        self.linever.setGeometry(QRect(310, 540, 180, 40))
        self.linever.setText('v2')
        
        self.linedir = QLineEdit(self.centralwidget)
        self.linedir.clear()
        self.linedir.setGeometry(QRect(310, 600, 380, 40))
        self.linedir.setText(r'D:\desktop\银行存款信息new.xlsx')

        self.textEdit = QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QRect(0, 0, 950, 150))
        self.textEdit.setObjectName("textEdit")
        
        self.lab1m = QLabel(self.centralwidget)
        self.lab1m.setText('1M')
        self.lab1m.setGeometry(QRect(10, 160, 40, 40))
        
        self.line1m = QLineEdit(self.centralwidget)
        self.line1m.clear()
        self.line1m.setGeometry(QRect(40, 160, 100, 40))
                        
        self.lab3m = QLabel(self.centralwidget)
        self.lab3m.setText('3M')
        self.lab3m.setGeometry(QRect(140, 160, 40, 40))
        
        self.line3m = QLineEdit(self.centralwidget)
        self.line3m.clear()
        self.line3m.setGeometry(QRect(170, 160, 100, 40))
        
        self.lab6m = QLabel(self.centralwidget)
        self.lab6m.setText('6M')
        self.lab6m.setGeometry(QRect(10, 220, 40, 40))
                
        self.line6m = QLineEdit(self.centralwidget)
        self.line6m.clear()
        self.line6m.setGeometry(QRect(40, 220, 100, 40))
                    
        self.lab9m = QLabel(self.centralwidget)
        self.lab9m.setText('9M')
        self.lab9m.setGeometry(QRect(140, 220, 40, 40))
        
        self.line9m = QLineEdit(self.centralwidget)
        self.line9m.clear()
        self.line9m.setGeometry(QRect(170, 220, 100, 40))
        
        self.lab1y = QLabel(self.centralwidget)
        self.lab1y.setText('1Y')
        self.lab1y.setGeometry(QRect(10, 280, 40, 40))
                
        self.line1y = QLineEdit(self.centralwidget)
        self.line1y.clear()
        self.line1y.setGeometry(QRect(40, 280, 100, 40))
                
        self.lab1mf = QLabel(self.centralwidget)
        self.lab1mf.setText('1M')
        self.lab1mf.setGeometry(QRect(10, 340, 40, 40))
        
        self.line1mf = QLineEdit(self.centralwidget)
        self.line1mf.clear()
        self.line1mf.setGeometry(QRect(40, 340, 100, 40))
                        
        self.lab3mf = QLabel(self.centralwidget)
        self.lab3mf.setText('3M')
        self.lab3mf.setGeometry(QRect(140, 340, 40, 40))
        
        self.line3mf = QLineEdit(self.centralwidget)
        self.line3mf.clear()
        self.line3mf.setGeometry(QRect(170, 340, 100, 40))
        
        self.lab6mf = QLabel(self.centralwidget)
        self.lab6mf.setText('6M')
        self.lab6mf.setGeometry(QRect(10, 400, 40, 40))
                
        self.line6mf = QLineEdit(self.centralwidget)
        self.line6mf.clear()
        self.line6mf.setGeometry(QRect(40, 400, 100, 40))
                    
        self.lab9mf = QLabel(self.centralwidget)
        self.lab9mf.setText('9M')
        self.lab9mf.setGeometry(QRect(140, 400, 40, 40))
        
        self.line9mf = QLineEdit(self.centralwidget)
        self.line9mf.clear()
        self.line9mf.setGeometry(QRect(170, 400, 100, 40))
        
        self.lab1yf = QLabel(self.centralwidget)
        self.lab1yf.setText('1Y')
        self.lab1yf.setGeometry(QRect(10, 460, 40, 40))
                
        self.line1yf = QLineEdit(self.centralwidget)
        self.line1yf.clear()
        self.line1yf.setGeometry(QRect(40, 460, 100, 40))
                
        self.labbank = QLabel(self.centralwidget)
        self.labbank.setText('银行名称')
        self.labbank.setGeometry(QRect(10, 500, 40, 40))
                
        self.linebank = QLineEdit(self.centralwidget)
        self.linebank.clear()
        self.linebank.setGeometry(QRect(20, 540, 150, 40))
        
        self.labbankf = QLabel(self.centralwidget)
        self.labbankf.setText('分行')
        self.labbankf.setGeometry(QRect(10, 560, 40, 40))
                
        self.linebankf = QLineEdit(self.centralwidget)
        self.linebankf.clear()
        self.linebankf.setGeometry(QRect(20, 600, 150, 40))
        
        self.labbei = QLabel(self.centralwidget)
        self.labbei.setText('备注')
        self.labbei.setGeometry(QRect(760, 155, 40, 40))
        
        self.beizhutext = QTextEdit(self.centralwidget)
        self.beizhutext.setGeometry(QRect(780, 190, 180, 210))
        self.beizhutext.setObjectName("textEdit")
        
        self.beizhutextm = QTextEdit(self.centralwidget)
        self.beizhutextm.setGeometry(QRect(780, 410, 180, 210))
        self.beizhutextm.setObjectName("textEdit")
        
        self.feichang = QTextEdit(self.centralwidget)
        self.feichang.setGeometry(QRect(500, 480, 190, 100))
        self.feichang.setObjectName("textEdit")

        self.feichanglab = QLabel(self.centralwidget)
        self.feichanglab.setText('非常规期限')
        self.feichanglab.setGeometry(QRect(500, 450, 150, 40))

        self.tablex = QTableView(self.centralwidget)
        self.modelx=QStandardItemModel(0,0)
        self.modelx.setRowCount(10)
        self.modelx.setColumnCount(2)
        
        self.modelx.setHorizontalHeaderLabels(['bank','分行','存款备注','存单备注'])        
        self.tablex.setModel(self.modelx)
        self.tablex.setGeometry(QRect(0, 690, 960, 200))
        self.tablex.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tablex.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        #self.tablex.setColumnWidth(0, 2)
        #self.tablex.setRowHeight(0,1)
                
        #item=QStandardItem('No') #一个QStandardItem就是一个单元格
        #self.modelx.setItem(0,0,item)

        #item.setTextAlignment(Qt.AlignCenter | Qt.AlignBottom)

        layout=QGridLayout()
        layout.addWidget(self.centralwidget,0,0)

        self.setLayout(layout)

        #将按钮的点击信号与相关的槽函数进行绑定，点击即触发
        self.pushButton.clicked.connect(self.btnPress1_clicked)
        self.pushButton_2.clicked.connect(self.btnPress2_clicked)
        self.pushButton_3.clicked.connect(self.naru)
        self.pushButton_4.clicked.connect(self.duqu)
        self.pushButton_5.clicked.connect(self.daochu)
        self.pushButton_6.clicked.connect(self.daochu2)
        
        self.df=pd.DataFrame(columns=[i for i in range(18)])
        
    def zhanshi(self):
        for i in range(len(self.tdf.index)):
            item=QStandardItem(str(self.tdf.iat[i,1])) #一个QStandardItem就是一个单元格
            self.modelx.setItem(i,0,item)
            item=QStandardItem(str(self.tdf.iat[i,2]))
            self.modelx.setItem(i,1,item)
            item=QStandardItem(str(self.tdf.iat[i,16]))
            self.modelx.setItem(i,2,item)    
            
            item=QStandardItem(str(self.tdf.iat[i,27]))
            self.modelx.setItem(i,3,item)

        

    def duqu(self):
        '''
        try:
            system('taskkill /F /IM excel.exe')
        except:
            pass
        '''
        self.labz.setText('正在导入……')
        QApplication.processEvents()
        
        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False
        app.screen_updating=False
        
        wb=app.books.open(self.linedir.text())
        hang=len(wb.sheets[0].range('A1').current_region.rows)
        lie=len(wb.sheets[0].range('A1').current_region.columns)
        self.df=wb.sheets[0].range((1,1),(hang,lie)).options(pd.DataFrame,index=False).value
  
        wb.close()
        
        #app.quit()
        app.kill()
        for i in range(5,len(self.df.index)-4):
            for j in [6,8,10,12,14,18,20,22,24]:
                if not(self.df.iat[i,j] is None):
                    self.df.iat[i,j]=None
        self.quchongflag=0
        hang=self.df.loc[self.df.iloc[:,0]=="可正式开展询价"].index[0]
        lie=self.df.loc[self.df.iloc[:,0]=="注释："].index[0]
        self.tdf=pd.DataFrame(self.df.iloc[hang:lie,:]).copy()
        self.fei=pd.Series(self.df.iloc[:,-1],index=self.df.index).copy()
        for i in range(len(self.tdf.index)):
            self.tdf.iat[i,-1]=None
        hang=pd.concat([self.df.iloc[1,:5],self.df.iloc[3,5:]],axis=0)
        self.tdf.columns=hang
        hang=self.tdf.iloc[:,5:].isnull().all(axis=1)
        hang=pd.Series([not(i) for i in hang],index=hang.index)
        hang=hang[hang]
        hang=list(hang.index)
        self.tdf=self.tdf.loc[hang].copy()
        self.df.columns=self.tdf.columns
        for i in range(len(self.fei)):
            if i>=5 and i<=len(self.fei)-4:
                self.fei.iat[i]="feikong" if i in hang else "kong"
            else:
                self.fei.iat[i]="feikong"
        self.tuse=pd.DataFrame('baidi',index=self.df.index,columns=self.df.columns)
        hang=self.df.iloc[5:,5:15].max(numeric_only=False)
        lie=self.df.iloc[5:,17:27].max(numeric_only=False)
        hang=list(hang)+[None,None]+list(lie)+[None,None]
        for i in range(5,len(self.df.columns)):
            for j in range(5,len(self.df.index)):
                if self.df.iat[j,i]==hang[i-5] and hang[i-5]!=None:
                    self.tuse.iat[j,i]='huangdi'
        self.quchongflag=[]        
        self.zhanshi()
        self.labz.setText('导入完成！')
        QApplication.processEvents()
        self.biaose=pd.DataFrame('heise',index=self.df.index,columns=self.df.columns)
        

    def btnPress1_clicked(self):
        zifu=self.textEdit.toPlainText()
        #第一步识别银行
        if zifu.find('南京银行')!=-1:
            self.linebank.setText('南京银行')
            if zifu.find('分行')!=-1:                
                self.linebankf.setText(zifu[zifu.find('分行')-2:zifu.find('分行')+2])
        elif zifu.find('兴业')!=-1:
            self.linebank.setText('兴业银行')
            if zifu.find('杭州')!=-1:                
                self.linebankf.setText('杭州分行')
        elif zifu.find('上银市南')!=-1:
            self.linebank.setText('上海银行')
            self.linebankf.setText('市南分行')
        elif zifu.find('杭州银行')!=-1:
            self.linebank.setText('杭州银行')
            self.linebankf.setText('上海分行')
        elif zifu.find('光大')!=-1:
            self.linebank.setText('光大银行')
            self.linebankf.setText('杭州分行')
        elif zifu.find('农行')!=-1:
            self.linebank.setText('农业银行')
            self.linebankf.setText('浙江省分行')
        elif zifu.find('浦发')!=-1:
            self.linebank.setText('浦发银行')
            self.linebankf.setText('上海分行')
            if zifu.find('广州')!=-1:                
                self.linebankf.setText('广州分行')
        elif zifu.find('华夏')!=-1:
            self.linebank.setText('华夏银行')
            self.linebankf.setText('上海分行')
            if zifu.find('杭分')!=-1:                
                self.linebankf.setText('杭州分行')
        elif zifu.find('广发上海')!=-1:
            self.linebank.setText('广发银行')
            self.linebankf.setText('上海分行')
        elif zifu.find('恒丰')!=-1:
            self.linebank.setText('恒丰银行')
            self.linebankf.setText('上海分行')
        elif zifu.find('浙商')!=-1:
            self.linebank.setText('浙商银行')
            self.linebankf.setText('杭州分行')
        elif zifu.find('中信')!=-1:
            self.linebank.setText('中信银行')
            self.linebankf.setText('杭州分行')
        elif zifu.find('桂林')!=-1:
            self.linebank.setText('桂林银行')
            self.linebankf.setText('总行')
        elif zifu.find('北京银行')!=-1:
            self.linebank.setText('北京银行')
            self.linebankf.setText('总行')
        elif zifu.find('交行')!=-1 or zifu.find('交通')!=-1:
            self.linebank.setText('交通银行')
            self.linebankf.setText('浙江分行')
        elif zifu.find('民生')!=-1:
            self.linebank.setText('民生银行')
            self.linebankf.setText('南京分行')
        elif zifu.find('华润')!=-1:
            self.linebank.setText('珠海华润银行')
            self.linebankf.setText('总行')
        elif zifu.find('中原')!=-1:
            self.linebank.setText('中原银行')
            self.linebankf.setText('三门峡分行')
        elif zifu.find('宁波')!=-1:
            self.linebank.setText('宁波通商银行')
            self.linebankf.setText('资金营运')
        
        #第二步，定存和存单割开
        hang='kongkong'
        lie='dangdang'
        if zifu.find('存单')!=-1:
            lie=zifu[zifu.find('存单'):]
            hang=zifu[:zifu.find(lie)]
        else:
            hang=zifu
        
        #第三步，识别利率
        
        if hang.find('1个月')!=-1:
            jiequ=self.sou(hang[hang.find('1个月')+2:])
            self.line1m.setText(jiequ)
        elif hang.find('1M')!=-1:
            jiequ=self.sou(hang[hang.find('1M')+2:])
            self.line1m.setText(jiequ)
        elif hang.find('1m')!=-1:
            jiequ=self.sou(hang[hang.find('1m')+2:])
            self.line1m.setText(jiequ)
        if hang.find('3个月')!=-1:
            jiequ=self.sou(hang[hang.find('3个月')+2:])
            self.line3m.setText(jiequ)
        elif hang.find('3M')!=-1:
            jiequ=self.sou(hang[hang.find('3M')+2:])
            self.line3m.setText(jiequ)
        elif hang.find('3m')!=-1:
            jiequ=self.sou(hang[hang.find('3m')+2:])
            self.line3m.setText(jiequ)
        if hang.find('6个月')!=-1:
            jiequ=self.sou(hang[hang.find('6个月')+2:])
            self.line6m.setText(jiequ)
        elif hang.find('6M')!=-1:
            jiequ=self.sou(hang[hang.find('6M')+2:])
            self.line6m.setText(jiequ)
        elif hang.find('6m')!=-1:
            jiequ=self.sou(hang[hang.find('6m')+2:])
            self.line6m.setText(jiequ)
        if hang.find('9个月')!=-1:
            jiequ=self.sou(hang[hang.find('9个月')+2:])
            self.line9m.setText(jiequ)
        elif hang.find('9M')!=-1:
            jiequ=self.sou(hang[hang.find('9M')+2:])
            self.line9m.setText(jiequ)
        elif hang.find('9m')!=-1:
            jiequ=self.sou(hang[hang.find('9m')+2:])
            self.line9m.setText(jiequ)
        if hang.find('1年')!=-1:
            jiequ=self.sou(hang[hang.find('1年')+2:])
            self.line1y.setText(jiequ)
        elif hang.find('1y')!=-1:
            jiequ=self.sou(hang[hang.find('1y')+2:])
            self.line1y.setText(jiequ)
        elif hang.find('1Y')!=-1:
            jiequ=self.sou(hang[hang.find('1Y')+2:])
            self.line1y.setText(jiequ)
            
        if lie.find('1个月')!=-1:
            jiequ=self.sou(lie[lie.find('1个月')+2:])
            self.line1mf.setText(jiequ)
        elif lie.find('1M')!=-1:
            jiequ=self.sou(lie[lie.find('1M')+2:])
            self.line1mf.setText(jiequ)
        elif lie.find('1m')!=-1:
            jiequ=self.sou(lie[lie.find('1m')+2:])
            self.line1mf.setText(jiequ)
        if lie.find('3个月')!=-1:
            jiequ=self.sou(lie[lie.find('3个月')+2:])
            self.line3mf.setText(jiequ)
        elif lie.find('3M')!=-1:
            jiequ=self.sou(lie[lie.find('3M')+2:])
            self.line3mf.setText(jiequ)
        elif lie.find('3m')!=-1:
            jiequ=self.sou(lie[lie.find('3m')+2:])
            self.line3mf.setText(jiequ)
        if lie.find('6个月')!=-1:
            jiequ=self.sou(lie[lie.find('6个月')+2:])
            self.line6mf.setText(jiequ)
        elif lie.find('6M')!=-1:
            jiequ=self.sou(lie[lie.find('6M')+2:])
            self.line6mf.setText(jiequ)
        elif lie.find('6m')!=-1:
            jiequ=self.sou(lie[lie.find('6m')+2:])
            self.line6mf.setText(jiequ)
        if lie.find('9个月')!=-1:
            jiequ=self.sou(lie[lie.find('9个月')+2:])
            self.line6mf.setText(jiequ)
        elif lie.find('9M')!=-1:
            jiequ=self.sou(lie[lie.find('9M')+2:])
            self.line9mf.setText(jiequ)
        elif lie.find('9m')!=-1:
            jiequ=self.sou(lie[lie.find('9m')+2:])
            self.line9mf.setText(jiequ)
        if lie.find('1年')!=-1:
            jiequ=self.sou(lie[lie.find('1年')+2:])
            self.line1yf.setText(jiequ)
        elif lie.find('1y')!=-1:
            jiequ=self.sou(lie[lie.find('1y')+2:])
            self.line1yf.setText(jiequ)
        elif lie.find('1Y')!=-1:
            jiequ=self.sou(lie[lie.find('1Y')+2:])
            self.line1yf.setText(jiequ)

        
        self.beizhutext.setPlainText(hang)
        self.beizhutextm.setPlainText(lie)
    
    def sou(self,ziduan):
        flag1=0
        flag2=0
        
        while flag1<=len(ziduan)-1:
            if '0123456789'.find(ziduan[flag1])==-1:
                flag1=flag1+1
            else:
                break
        flag2=flag1
        while flag2<=len(ziduan)-1:
            if '0123456789'.find(ziduan[flag2])!=-1 or '.'.find(ziduan[flag2])!=-1:
                flag2=flag2+1
            else:
                break
        
        return ziduan[flag1:flag2]
                
    

    def btnPress2_clicked(self):
        self.textEdit.clear()
        self.line1m.clear()
        self.line3m.clear()
        self.line6m.clear()
        self.line9m.clear()
        self.line1y.clear()
        self.line1mf.clear()
        self.line3mf.clear()
        self.line6mf.clear()
        self.line9mf.clear()
        self.line1yf.clear()
        self.linebank.clear()
        self.linebankf.clear()
        self.beizhutext.clear()
        self.beizhutextm.clear()
        self.feichang.clear()

        
    
    def naru(self):
        new=pd.DataFrame([None for i in range(len(self.tdf.columns))]).T
        new.columns=self.tdf.columns
        new.index=[500]
        new.iat[0,1]=self.linebank.text() if self.linebank.text()!='' else None
        new.iat[0,2]=self.linebankf.text() if self.linebankf.text()!='' else None
        new.iat[0,5]=float(self.line1m.text())/100 if self.line1m.text()!='' else None
        new.iat[0,7]=float(self.line3m.text())/100 if self.line3m.text()!='' else None
        new.iat[0,9]=float(self.line6m.text())/100 if self.line6m.text()!='' else None
        new.iat[0,11]=float(self.line9m.text())/100 if self.line9m.text()!='' else None
        new.iat[0,13]=float(self.line1y.text())/100 if self.line1y.text()!='' else None
        new.iat[0,15]=self.feichang.toPlainText()
        new.iat[0,16]=self.beizhutext.toPlainText()
        new.iat[0,17]=float(self.line1mf.text())/100 if self.line1mf.text()!='' else None
        new.iat[0,19]=float(self.line3mf.text())/100 if self.line3mf.text()!='' else None
        new.iat[0,21]=float(self.line6mf.text())/100 if self.line6mf.text()!='' else None
        new.iat[0,23]=float(self.line9mf.text())/100 if self.line9mf.text()!='' else None
        new.iat[0,25]=float(self.line1yf.text())/100 if self.line1yf.text()!='' else None
        new.iat[0,27]=self.beizhutextm.toPlainText()

        newstr=str(new.iat[0,1])+str(new.iat[0,2])
        dfstr=[str(self.df.iat[i,1])+str(self.df.iat[i,2]) for i in range(len(self.df.index))]

        if (newstr in dfstr)and(new.iat[0,1]!=None)and(new.iat[0,2]!=None):
            self.tdf=pd.concat([self.tdf,new],axis=0)
            self.quchong() 
            self.labz.setText('已纳入！')
            QApplication.processEvents()
        else:
            str1=''
            for k in self.df.index[:-4]:
                if (new['银行'][500][0] in str(self.df['银行'][k])):
                    str1=str1+str(self.df['银行'][k])+str(self.df['分行'][k])+' '
  
            if str1=='':
                self.labz.setText('银行和分行输入错误!且在数据库中也没找到相关银行（以“银行”栏内第一个汉字为准进行搜索）')
            else:
                self.labz.setText('银行和分行输入错误！您要的是不是'+str1+'?')
            QApplication.processEvents()
        
        for i in range(len(self.tdf.index)):
            if (self.tdf.iat[i,16] is None) or (self.tdf.iat[i,16]==""):
                self.tdf.iat[i,16]=self.df.loc[self.tdf.index[i]][16]
            else:
                self.df.loc[self.tdf.index[i]][16]=self.tdf.iat[i,16]
            if (self.tdf.iat[i,27] is None) or (self.tdf.iat[i,27]==""):
                self.tdf.iat[i,27]=self.df.loc[self.tdf.index[i]][27]
            else:
                self.df.loc[self.tdf.index[i]][27]=self.tdf.iat[i,27]
        self.zhanshi()
        


    
    def quchong(self):
        #去重并标色
        hang=self.tdf[(self.tdf['银行']==self.tdf.iat[-1,1])]
        hang=hang[hang['分行']==self.tdf.iat[-1,2]]
        self.quchongflag.append(0)

        if len(hang.index)==2:
            for i in range(0,5):
                self.tdf.iat[-1,i]=hang.iat[0,i]
            self.tdf.drop(inplace=True,index=hang.index[0])
            yz=list(self.tdf.index)
            yz[-1]=hang.index[0]
            self.tdf.index=yz
            self.tdf.sort_index(ascending=True,inplace=True,axis=0)
        else:
            lie=self.df[self.df.iloc[:,1]==hang.iat[0,1]]
            lie=lie[lie.iloc[:,2]==hang.iat[0,2]]
            for i in range(0,5):
                self.tdf.iat[-1,i]=self.df.iat[lie.index[0],i]
            yz=list(self.tdf.index)
            yz[-1]=lie.index[0]
            self.quchongflag[-1]=lie.index[0]
            #self.biaose.iloc[lie.index[0],1:]=['lanse' for m in range(len(self.biaose.columns)-1)]
            self.tdf.index=yz
            self.tdf.sort_index(ascending=True,inplace=True,axis=0)
        
    def daochu(self):
        hang=list(self.tdf.index)
        self.biaose=pd.DataFrame('heise',index=self.df.index,columns=self.df.columns)
        for k in range(len(self.quchongflag)):
            if self.quchongflag[k]!=0:
                for i in range(1,len(self.biaose.columns)):
                    self.biaose.iat[self.quchongflag[k],i]='lanse'
        
        for i in range(len(hang)):
            for j in range(5,len(self.df.columns)):  
                
                if j!=16 and j!=28 and j!=15 and j!=27:
                    if self.df.iat[hang[i],j]!=None and self.tdf.iat[i,j]!=None:
                        if float(self.df.iat[hang[i],j])<float(self.tdf.iat[i,j]):
                            self.biaose.iat[hang[i],j]='red'
                        elif float(self.df.iat[hang[i],j])>float(self.tdf.iat[i,j]):
                            self.biaose.iat[hang[i],j]='green'
                        self.df.iat[hang[i],j]=self.tdf.iat[i,j]
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]==None:
                        pass
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]!=None:
                        self.biaose.iat[hang[i],j]='lanse'
                        self.biaose.iat[hang[i],28]='lanse'
                        self.df.iat[hang[i],j]=self.tdf.iat[i,j]
                        
                    
                elif j==16 or j==15 or j==27:
                    self.biaose.iat[hang[i],j]='lanse'
                    '''
                    if self.df.iat[hang[i],j]!=None and self.tdf.iat[i,j]!=None and self.df.iat[hang[i],j]!=self.tdf.iat[i,j]:
                        self.biaose.iat[hang[i],j]='lanse'
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]==None:
                        pass
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]!=None:
                        self.biaose.iat[hang[i],j]='lanse'
                    self.df.iat[hang[i],j]=self.tdf.iat[i,j]
                    '''

        xtuse=pd.DataFrame('baidi',index=self.df.index,columns=self.df.columns)
        hang=self.df.iloc[5:,5:15].max(numeric_only=False)
        lie=self.df.iloc[5:,17:27].max(numeric_only=False)
        hang=list(hang)+[None,None]+list(lie)+[None,None]
        for i in range(5,len(self.df.columns)):
            for j in range(5,len(self.df.index)):
                if self.df.iat[j,i]==hang[i-5] and hang[i-5]!=None:
                    xtuse.iat[j,i]='huangdi'
      
        self.labz.setText('正在导出……')
        QApplication.processEvents()


        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False
        app.screen_updating=False
        
        wb=app.books.open(self.linedir.text())
        
        lie=''

        for i in range(len(self.df.index)):
        #for i in range(5,6):
            if i>=4 and i<=len(self.df.index)-5:
                hang=list(self.df.iloc[i,5:28])
                if all([k is None for k in hang]):
                    lie='kong'
                else:
                    lie='feikong'
                if lie!=self.fei[i]:
                    self.fei[i]=lie
                wb.sheets[0].range((i+2,29)).value=lie
                for j in range(len(self.df.columns)):
                    
                    if (self.biaose.iat[i,j]=='lanse' or self.biaose.iat[i,j]=='green' or self.biaose.iat[i,j]=='red'):
                        if j>=5 and j<=27:
                            wb.sheets[0].range((i+2,j+1)).value=self.df.iat[i,j]
                        elif j==28:
                            wb.sheets[0].range((i+2,j+1)).value=self.fei[i]
                            wb.sheets[0].range((i+2,j+1)).api.Font.Color = 0xFFFFFF
                    
                    '''
                    if (wb.sheets[0].range((i+2,j+1)).api.Font.Color!=0x000000)and(j>=5)and(j<=27):
                        if self.biaose.iat[i,j]=='heise':
                            wb.sheets[0].range((i+2,j+1)).api.Font.Color = 0x000000
                    '''
                    
                    if self.biaose.iat[i,j]=='red' and j<=27:
                        wb.sheets[0].range((i+2,j+1)).api.Font.ColorIndex = 3
                    elif self.biaose.iat[i,j]=='lanse' and j<=27 and not(j in [15,16,27]):
                        wb.sheets[0].range((i+2,j+1)).api.Font.ColorIndex = 5
                    elif self.biaose.iat[i,j]=='green' and j<=27:
                        wb.sheets[0].range((i+2,j+1)).api.Font.Color = 0x00FF00

                    if xtuse.iat[i,j]=='huangdi':
                        wb.sheets[0].range((i+2,j+1)).color = (250,218,141)
                        wb.sheets[0].range((i+2,j+2)).color = (250,218,141)

       
        wb.sheets[0].range(2,1).value=self.lineday.text()
        wb.sheets[0].range(2,2).value=self.linever.text()
        wb.sheets[0].range((2,1),(57,29)).api.AutoFilter(field:=int(29), Criteria1:="feikong", True)
        wb.save(self.linedir.text())

        
        wb.close()
          
        #app.quit()
        app.kill()
        self.labz.setText('导出已完成')

    def daochu2(self):
        hang=list(self.tdf.index)
        self.biaose=pd.DataFrame('heise',index=self.df.index,columns=self.df.columns)
        for k in range(len(self.quchongflag)):
            if self.quchongflag[k]!=0:
                for i in range(1,len(self.biaose.columns)):
                    self.biaose.iat[self.quchongflag[k],i]='lanse'
        
        for i in range(len(hang)):
            for j in range(5,len(self.df.columns)):  
                
                if j!=16 and j!=28 and j!=15 and j!=27:
                    if self.df.iat[hang[i],j]!=None and self.tdf.iat[i,j]!=None:
                        if float(self.df.iat[hang[i],j])<float(self.tdf.iat[i,j]):
                            self.biaose.iat[hang[i],j]='red'
                        elif float(self.df.iat[hang[i],j])>float(self.tdf.iat[i,j]):
                            self.biaose.iat[hang[i],j]='green'
                        self.df.iat[hang[i],j]=self.tdf.iat[i,j]
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]==None:
                        pass
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]!=None:
                        self.biaose.iat[hang[i],j]='lanse'
                        self.biaose.iat[hang[i],28]='lanse'
                        self.df.iat[hang[i],j]=self.tdf.iat[i,j]
                        
                    
                elif j==16 or j==15 or j==27:
                    
                    self.biaose.iat[hang[i],j]='lanse'
                    '''
                    if self.df.iat[hang[i],j]!=None and self.tdf.iat[i,j]!=None and self.df.iat[hang[i],j]!=self.tdf.iat[i,j]:
                        self.biaose.iat[hang[i],j]='lanse'
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]==None:
                        pass
                    elif self.df.iat[hang[i],j]==None and self.tdf.iat[i,j]!=None:
                        self.biaose.iat[hang[i],j]='lanse'
                    self.df.iat[hang[i],j]=self.tdf.iat[i,j]
                    '''

        xtuse=pd.DataFrame('baidi',index=self.df.index,columns=self.df.columns)
        hang=self.df.iloc[5:,5:15].max(numeric_only=False)
        lie=self.df.iloc[5:,17:27].max(numeric_only=False)
        hang=list(hang)+[None,None]+list(lie)+[None,None]
        for i in range(5,len(self.df.columns)):
            for j in range(5,len(self.df.index)):
                if self.df.iat[j,i]==hang[i-5] and hang[i-5]!=None:
                    xtuse.iat[j,i]='huangdi'
            
      
        self.labz.setText('正在导出……')
        QApplication.processEvents()


        app=xw.App(visible=False,add_book=False)
        app.display_alerts=False
        app.screen_updating=False
        
        wb=app.books.open(self.linedir.text())
        
        lie=''

        for i in range(len(self.df.index)):
        #for i in range(5,6):
            self.labz.setText('正在导出……目前第'+str(i)+'行')
            QApplication.processEvents()
            if i>=4 and i<=len(self.df.index)-5:
                hang=list(self.df.iloc[i,5:28])
                if all([k is None for k in hang]):
                    lie='kong'
                else:
                    lie='feikong'
                if lie!=self.fei[i]:
                    self.fei[i]=lie
                wb.sheets[0].range((i+2,29)).value=lie
                for m in [6,8,10,12,14,18,20,22,24]:
                    wb.sheets[0].range((i+2,m+1)).value=None
                for j in range(len(self.df.columns)):
                    
                    if (self.biaose.iat[i,j]=='lanse' or self.biaose.iat[i,j]=='green' or self.biaose.iat[i,j]=='red'):
                        if j>=5 and j<=27:
                            wb.sheets[0].range((i+2,j+1)).value=self.df.iat[i,j]
                        elif j==28:
                            wb.sheets[0].range((i+2,j+1)).value=self.fei[i]
                            #wb.sheets[0].range((i+2,j+1)).api.Font.Color = 0xFFFFFF
                    
                    
                    if (j<=27):
                        wb.sheets[0].range((i+2,j+1)).api.Font.Color = 0x000000
                    elif j==28:
                        wb.sheets[0].range((i+2,j+1)).api.Font.Color = 0xFFFFFF
                    
                    if wb.sheets[0].range((i+2,j+1)).color != (255,255,255):
                        wb.sheets[0].range((i+2,j+1)).color = (255,255,255)

       
        wb.sheets[0].range(2,1).value=self.lineday.text()
        wb.sheets[0].range(2,2).value=self.linever.text()
        wb.sheets[0].range((2,1),(57,29)).api.AutoFilter(field:=int(29), Criteria1:="feikong", True)
        
        wb.save(self.linedir.text())

        
        #wb.save(self.linedir.text())

        
        wb.close()
          
        #app.quit()
        app.kill()
        self.labz.setText('导出已完成')


    
if __name__ == '__main__':
    app=QApplication(sys.argv)
    win=TextEditDemo()
    win.show()
    sys.exit(app.exec_())
    












