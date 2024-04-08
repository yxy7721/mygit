# -*- coding: utf-8 -*-
"""
Created on Tue Jun 27 14:12:59 2023

@author: yangxy
"""

import pandas as pd, numpy as np, time

data= pd.read_csv(r"D:\desktop\learn_LGBM\flights_10k.csv")

# 提取有用的列
data= data[["MONTH","DAY","DAY_OF_WEEK","AIRLINE","FLIGHT_NUMBER","DESTINATION_AIRPORT",
                 "ORIGIN_AIRPORT","AIR_TIME", "DEPARTURE_TIME","DISTANCE","ARRIVAL_DELAY"]]
data.dropna(inplace=True)

from sklearn.model_selection import train_test_split
# 筛选出部分数据
data["ARRIVAL_DELAY"] = (data["ARRIVAL_DELAY"]>10)*1

# 以下四列数据转换为类别
cols = ["AIRLINE","FLIGHT_NUMBER","DESTINATION_AIRPORT","ORIGIN_AIRPORT"]
for item in cols:
    data[item] = data[item].astype("category").cat.codes +1
del item,cols
# 划分训练集和测试集
train, test, y_train, y_test = train_test_split(
                                    data.drop(["ARRIVAL_DELAY"], axis=1), 
                                    data["ARRIVAL_DELAY"], 
                                    random_state=10, test_size=0.25)
data


from matplotlib import pyplot as plt
from sklearn.metrics import accuracy_score,roc_auc_score
import lightgbm as lgb
from lightgbm import LGBMClassifier

#人肉调maxdepth参数
def test_depth(max_depth):    
    gbm = lgb.LGBMClassifier(max_depth=max_depth)
    gbm.fit(train, y_train,
            eval_set=[(test, y_test)],
            eval_metric='binary_logloss',
            callbacks=[lgb.early_stopping(5)])
    #eval_metric默认值：LGBMRegressor 为“l2”，LGBMClassifier 为“logloss”，LGBMRanker 为“ndcg”。
    #使用binary_logloss或者logloss准确率都是一样的。默认logloss
    y_pred = gbm.predict(test)
    # 计算准确率
    accuracy = accuracy_score(y_test,y_pred)
    auc_score=roc_auc_score(y_test,gbm.predict_proba(test)[:,1])#predict_proba输出正负样本概率值，取第二列为正样本概率值
    print("max_depth=",max_depth,"accuarcy: %.2f%%" % (accuracy*100.0),"auc_score: %.2f%%" % (auc_score*100.0))

test_depth(3)
test_depth(5)
test_depth(6)
test_depth(9)

#网格调参数
parameters = {
              'max_depth': [8, 10, 12],
              'learning_rate': [0.05, 0.1, 0.15],
              'n_estimators': [100, 200,500],
              "num_leaves":[25,31,36]}

gbm = lgb.LGBMClassifier(max_depth=10,#构建树的深度，越大越容易过拟合
            learning_rate=0.01,
            n_estimators=100,         
            seed=0,
            missing=None)

from sklearn.model_selection import GridSearchCV 
gs = GridSearchCV(gbm, param_grid=parameters, scoring='accuracy', cv=3)
gs.fit(train, y_train)

print("Best score: %0.3f" % gs.best_score_)
print("Best parameters set: %s" % gs.best_params_ )
y_pred = gs.predict(test)

# 计算准确率和auc
accuracy = accuracy_score(y_test,y_pred)
auc_score=roc_auc_score(y_test,gs.predict_proba(test)[:,1])#predict_proba输出正负样本概率值，取第二列为正样本概率值
print("accuarcy: %.2f%%" % (accuracy*100.0),"auc_score: %.2f%%" % (auc_score*100.0))











