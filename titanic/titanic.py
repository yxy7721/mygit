# -*- coding: utf-8 -*-
"""
Created on Mon Jun 26 09:42:26 2023

@author: yangxy
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from category_encoders import TargetEncoder#需要numpy版本早于1.23
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.metrics import confusion_matrix,f1_score,accuracy_score
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis as LDA
from imblearn.metrics import geometric_mean_score
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import OneHotEncoder
from sklearn.feature_selection import SelectFromModel
from sklearn.model_selection import cross_val_score
from sklearn.metrics import roc_curve,auc,classification_report,mean_squared_error
from sklearn.svm import SVC
import toad
import lightgbm as lgb
from lightgbm import LGBMClassifier

train = pd.read_csv(r"D:/desktop/titanic/train.csv")
test = pd.read_csv(r"D:/desktop/titanic/test.csv")

(train.isna().sum()/train.shape[0]).apply(lambda x:format(x,'.2%'))

train.info()
train.select_dtypes(include="object")

#生成称呼
train_process = train.set_index(['PassengerId'])
train_process = train_process.drop(['Cabin'],axis=1)
train_process['Called'] = train_process['Name'].str.findall('Miss|Mr|Ms').str[0].to_frame()
train_process['Name_length'] = train_process['Name'].apply(lambda x:len(x))
train_process['First_name'] = train_process['Name'].str.split(',').str[0]
train_process = train_process.drop(['Name'],axis=1)
train_process

test_process = test.set_index(['PassengerId'])
test_process = test_process.drop(['Cabin'],axis=1)
test_process['Called'] = test_process['Name'].str.findall('Miss|Mr|Ms').str[0].to_frame()
test_process['Name_length'] = test_process['Name'].apply(lambda x:len(x))
test_process['First_name'] = test_process['Name'].str.split(',').str[0]
test_process = test_process.drop(['Name'],axis=1)
test_process


#生成target编码
X_train = train_process.iloc[:,1:]
y_train = train_process.iloc[:,0]
X_test = test_process
tar_encoder1 = TargetEncoder(cols=['Sex','Ticket','Embarked','Called','Name_length','First_name'],
                             handle_missing='value',
                             handle_unknown='value')
tar_encoder1.fit(X_train,y_train)
X_train_encode = tar_encoder1.transform(X_train)
X_test_encode = tar_encoder1.transform(X_test)
X_train_encode.isna().sum()
X_train.isna().sum()

del tar_encoder1,test,test_process,train,train_process,X_test,X_train
 
#分箱
train_binning = pd.concat([X_train_encode[['Age','Fare']],y_train],axis=1)
test_binning = pd.concat([X_train_encode[['Age','Fare']]],axis=1)
c_tree = toad.transform.Combiner()
c_tree.fit(train_binning,y='Survived',method='dt',min_samples=0.05,n_bins=6)

train_binned = c_tree.transform(pd.concat([X_train_encode[['Age','Fare']],y_train],axis=1),labels=False)
test_binned = c_tree.transform(pd.concat([X_test_encode[['Age','Fare']]],axis=1),labels=False)

tar_encoder2 = TargetEncoder(cols=['Age','Fare'],
                             handle_missing='value',
                             handle_unknown='value')
tar_encoder2.fit(train_binned.iloc[:,:-1],train_binned.iloc[:,-1])
train_binned_target = tar_encoder2.transform(train_binned.iloc[:,:-1])
test_binned_target = tar_encoder2.transform(test_binned)
X_train_final = pd.concat([X_train_encode.drop(['Age','Fare'],axis=1),train_binned_target],axis=1)
X_test_final = pd.concat([X_test_encode.drop(['Age','Fare'],axis=1),test_binned_target],axis=1)

del c_tree,tar_encoder2,test_binned,test_binned_target,test_binning
del train_binned,train_binned_target,train_binning,X_test_encode,X_train_encode

#不平衡处理
from imblearn.over_sampling import SMOTE
smo = SMOTE(sampling_strategy='minority')
X_smo,y_smo = smo.fit_resample(X_train_final,y_train)

#随机森林找参数，先找森林个数这个参数
estimator_score = {}
for est in np.arange(30,500,10):
    score = cross_val_score(estimator=RandomForestClassifier(n_estimators=est),X=X_smo,
                            y=y_smo,cv=10,n_jobs=-1,scoring='f1').mean()
    estimator_score[est] = score

pd.DataFrame([estimator_score],index=['n_estimators']).T.sort_values(by='n_estimators',ascending=False)
#其他参数类似

#用随机森林正式预测
rf_clf_final = RandomForestClassifier(n_estimators=180,
                                      max_depth=8,
                                      min_samples_split=8,
                                      min_samples_leaf=2)
rf_clf_final.fit(X_smo,y_smo)
score=cross_val_score(rf_clf_final,
                      X =X_smo,
                      y=y_smo,
                      verbose=0)
y_pre = rf_clf_final.predict(X_test_final)
real_label = pd.read_csv(r'D:/desktop/titanic/gender_submission.csv')
y_real = real_label['Survived'].values
confusion_matrix(y_real,y_pre)
f1_score(y_real,y_pre)
classification_report(y_real,y_pre)
mean_squared_error(y_real,y_pre)
y_real.mean()

y_pre=pd.DataFrame(y_pre)
y_pre["PassengerId"]=X_test_final.index
y_pre=y_pre.set_index("PassengerId")
y_pre.columns=[real_label.columns[1]]
y_pre.to_csv(r'D:/desktop/gender_submission.csv',header=True,index=True)

#用lightgbm的算法
gbm = lgb.LGBMClassifier()
gbm.fit(X_smo,y_smo)

y_pre = gbm.predict(X_test_final)
real_label = pd.read_csv(r'D:/desktop/titanic/gender_submission.csv')
y_real = real_label['Survived'].values
import graphviz
lgb.create_tree_digraph(gbm, tree_index=1)



y_pre=pd.DataFrame(y_pre)
y_pre["PassengerId"]=X_test_final.index
y_pre=y_pre.set_index("PassengerId")
y_pre.columns=[real_label.columns[1]]
y_pre.to_csv(r'D:/desktop/gender_submission.csv',header=True,index=True)







