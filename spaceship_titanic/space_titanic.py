# -*- coding: utf-8 -*-
"""
Created on Wed Jun 28 13:08:53 2023

@author: yangxy
"""

# data analysis and wrangling
import pandas as pd
import numpy as np

# visualization
import seaborn as sns
import matplotlib.pyplot as plt

#other
import copy

#ML
from sklearn.preprocessing import OneHotEncoder
from sklearn.preprocessing import KBinsDiscretizer

train_df = pd.read_csv(r'D:/desktop/spaceship_titanic/train.csv')
test_df = pd.read_csv(r'D:/desktop/spaceship_titanic/test.csv')

len(train_df.index)
print(train_df.columns)

train_df.loc[:,['RoomService', 'FoodCourt', 'ShoppingMall', 'Spa', 'VRDeck']]
train_df.info()

train_df[['Age', 'Transported']].groupby(
                        ['Age'], as_index=False).mean(
                            ).sort_values(by='Transported', ascending=False
                         )
                                          
#sns.countplot(x="CryoSleep",hue="Transported",data=train_df)

#sns.countplot(x="Transported",data=train_df)

train_df["PassengerId"].describe()

#处理PassengerId
def exe(df):
    #df=copy.deepcopy(train_df)
    df["PId1"]=df["PassengerId"].str.split("_",expand=True)[0]
    df["PId2"]=df["PassengerId"].str.split("_",expand=True)[1]
    df.drop(["PassengerId"],axis=1,inplace=True)
    
    df["PId1"]=pd.to_numeric(df["PId1"],errors="coerce")
    df["PId2"]=pd.to_numeric(df["PId2"],errors="coerce")
    
    df2=df[["PId1","PId2"]]
    df2=df2.groupby("PId1", as_index=False).max()
    
    df["PId2Max"]=df["PId1"].map(lambda x:df2[df2["PId1"]==x]["PId2"].iat[0])

    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理HomePlanet
#sns.countplot(x="HomePlanet",hue="Transported",data=train_df)
#sns.countplot(x="Transported",data=train_df)

train_df["HomePlanet"].unique()
train_df["HomePlanet"].isna().sum()

def exe(df):
    #df=copy.deepcopy(train_df)
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[["HomePlanet"]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:"Home_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=["HomePlanet"],axis=1,inplace=True)
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))


#处理CroyoSleep
train_df.columns
df=copy.deepcopy(train_df)

df["CryoSleep"].isna().sum()

#sns.countplot(x="CryoSleep",hue="Transported",data=train_df)

def exe(df):
    #df=copy.deepcopy(train_df)
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[["CryoSleep"]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:"Sleep_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=["CryoSleep"],axis=1,inplace=True)
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理Cabin
def exe(df):
    #df=copy.deepcopy(train_df)    
    newcabin=df["Cabin"].str.split("/",expand=True)
    newcabin.columns=pd.Series(["A","B","C"]).map(lambda x:"Cabin"+x)
    df=pd.concat([df,newcabin],axis=1)
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#sns.countplot(x="CabinA",hue="Transported",data=train_df)

def exe(df,collist):
    #df=copy.deepcopy(train_df)
    #col="CabinA"
    for col in collist:
        onehot_encoder = OneHotEncoder(sparse_output=False)
        onehot_encoded = onehot_encoder.fit_transform(df[[col]])
        onehot_encoded=pd.DataFrame(onehot_encoded)
        onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
        df=pd.concat([df,onehot_encoded],axis=1)
        df.drop(labels=[col],axis=1,inplace=True)
    return df

train_df=copy.deepcopy(exe(train_df,["CabinA"]))
test_df=copy.deepcopy(exe(test_df,["CabinA"]))

def exe(df):
    df["CabinB"]=pd.to_numeric(df["CabinB"],errors="ignore")
    meantmp=df["CabinB"].mean(skipna=True)
    df["CabinB"]=df["CabinB"].fillna(value=meantmp)
    train_df=copy.deepcopy(df)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

def exe(df):
    #df=copy.deepcopy(train_df)
    col="CabinC"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    del df["Cabin"]
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理destination
#sns.countplot(x="CabinB",hue="Transported",data=df)

def exe(df):
    #df=copy.deepcopy(train_df)
    col="Destination"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))


#处理age
def exe(df):
    #df=copy.deepcopy(train_df)
    #sns.distplot(df["Age"],rug=False)
    #sns.stripplot(data=df,y="Age",x="Transported")
    df["Age"]=df["Age"].fillna(df["Age"].mean(skipna=True))
    df["Age"].isna().sum()
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理VIP
df=copy.deepcopy(train_df)
df.columns
def exe(df):
    #df=copy.deepcopy(train_df)
    col="VIP"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#五种消费联动处理
df=copy.deepcopy(train_df)
df.columns
def exe(df):
    tmp1=copy.deepcopy(
        df.loc[:,['RoomService', 'FoodCourt', 'ShoppingMall', 'Spa', 'VRDeck']]
        )
    tmp1.isna().sum()
    df["MostSpent"]=tmp1.idxmax(axis=1)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

df=copy.deepcopy(train_df)
df.columns
def exe(df):
    tmp1=copy.deepcopy(
        df.loc[:,['RoomService', 'FoodCourt', 'ShoppingMall', 'Spa', 'VRDeck']]
        )
    df["AveSpent"]=tmp1.mean(axis=1)
    df["AveSpent"].isna().sum()
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))



#处理RoomServeive
def exe(df):
    #df=copy.deepcopy(train_df)
    df.columns
    df["RoomService"].unique()
    #sns.distplot(df["RoomService"],rug=False)
    kbd=KBinsDiscretizer(n_bins=100,encode="onehot",strategy="quantile")
    tmp1=df["RoomService"].fillna(df["RoomService"].mean())
    tmp1.isna().sum()
    tmp1=np.array(tmp1).reshape(-1,1)
    tmp1=kbd.fit_transform(tmp1).todense()
    tmp1=pd.DataFrame(tmp1)
    tmp1.columns=tmp1.columns.map(lambda x:"RoomService"+"_"+str(x))
    df=pd.concat([df,tmp1],axis=1)
    df=df.drop(["RoomService"],axis=1)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理FoodCourt
df=copy.deepcopy(train_df)
df.columns
df["FoodCourt"].unique()
def exe(df):
    #sns.distplot(df["FoodCourt"],rug=False)
    tmp1=copy.deepcopy(df["FoodCourt"])
    df.loc[tmp1==0,"FoodC"]="zero"
    df.loc[(tmp1>0) & (tmp1<=500),"FoodC"]="small"
    df.loc[(tmp1>500),"FoodC"]="big"
    #sns.countplot(df["FoodC"])
    df=df.drop(["FoodCourt"],axis=1)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

def exe(df):
    #df=copy.deepcopy(train_df)
    col="FoodC"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理shoppingmall
df=copy.deepcopy(train_df)
df.columns
df["ShoppingMall"].unique()
#sns.distplot(df["ShoppingMall"],rug=False)
def exe(df):
    tmp1=copy.deepcopy(df["ShoppingMall"])
    df.loc[tmp1==0,"shop"]="zero"
    df.loc[(tmp1>0) & (tmp1<=500),"shop"]="small"
    df.loc[(tmp1>500),"shop"]="big"
    #sns.countplot(df["shop"])
    df.drop(["ShoppingMall"],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

def exe(df):
    #df=copy.deepcopy(train_df)
    col="shop"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df

train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理spa
df=copy.deepcopy(train_df)
df.columns
df["Spa"].unique()
#sns.distplot(df["Spa"],rug=False)
def exe(df):
    tmp1=copy.deepcopy(df["Spa"])
    df.loc[tmp1==0,"sp"]="zero"
    df.loc[(tmp1>0) & (tmp1<=200),"sp"]="small"
    df.loc[(tmp1>200),"sp"]="big"
    #sns.countplot(df["sp"])
    df.drop(["Spa"],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

def exe(df):
    #df=copy.deepcopy(train_df)
    col="sp"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理VRdeck
df=copy.deepcopy(train_df)
df.columns
df["VRDeck"].unique()
#sns.distplot(df["VRDeck"],rug=False)
def exe(df):
    tmp1=copy.deepcopy(df["VRDeck"])
    df.loc[tmp1==0,"vr"]="zero"
    df.loc[(tmp1>0) & (tmp1<=200),"vr"]="small"
    df.loc[(tmp1>200),"vr"]="big"
    #sns.countplot(df["vr"])
    df.drop(["VRDeck"],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

def exe(df):
    #df=copy.deepcopy(train_df)
    col="vr"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理Name
def exe(df):
    #df=copy.deepcopy(train_df)
    #df.columns
    tmp1=df["Name"].str.split(" ",expand=True)
    tmp1.columns=["fname","sname"]
    df["sname"]=tmp1["sname"]
    len(df["sname"].unique())
    df["sname"].fillna(0,inplace=True)
    tmp1=df.groupby(by="sname").count()["Age"]
    tmp1=tmp1.to_dict()
    df["sname"].isna().sum()
    df["Nsname"]=df["sname"].map(lambda x:tmp1[x])
    df.drop(["sname","Name"],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理Transported
df=copy.deepcopy(train_df)
df.columns
def exe(df):
    df["Transported"].isna().sum()
    df["Transported"]=df["Transported"]*1
    return df
train_df=copy.deepcopy(exe(train_df))
#test_df=copy.deepcopy(exe(test_df))

#处理MostSpent
df=copy.deepcopy(train_df)
df.columns
def exe(df):
    #df=copy.deepcopy(train_df)
    col="MostSpent"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#处理AveSpent
df=copy.deepcopy(train_df)
df.columns
def exe(df):
    pd.Series(df["AveSpent"].unique()).sort_values()
    tmp1=pd.cut(df["AveSpent"],bins=[-1,50,500,5000,10000])
    #sns.countplot(tmp1)
    df["AveSpent"]=tmp1
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

def exe(df):
    #df=copy.deepcopy(train_df)
    col="AveSpent"
    onehot_encoder = OneHotEncoder(sparse_output=False)
    onehot_encoded = onehot_encoder.fit_transform(df[[col]])
    onehot_encoded=pd.DataFrame(onehot_encoded)
    onehot_encoded.columns=onehot_encoded.columns.map(lambda x:col+"_"+str(x))
    df=pd.concat([df,onehot_encoded],axis=1)
    df.drop(labels=[col],axis=1,inplace=True)
    return df
train_df=copy.deepcopy(exe(train_df))
test_df=copy.deepcopy(exe(test_df))

#ML过程
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.ensemble import GradientBoostingClassifier

X_train, X_test, y_train, y_test = train_test_split(
    train_df.drop(['Transported'],axis=1), 
    train_df['Transported'], 
    random_state=42
    )

#用随机森林
def random_forest():
    rfc=RandomForestClassifier(random_state=1)
    rfc.fit(X_train, y_train)
    
    rfc.score(X_test,y_test)
    tmp1=pd.Series(rfc.feature_importances_,index=train_df.drop(['Transported'],axis=1).columns)
    tmp1=tmp1.sort_values(ascending=False)

    origin_test_df = pd.read_csv(r'D:/desktop/spaceship_titanic/test.csv')
    origin_test_df["PassengerId"]
    outcome=rfc.predict(test_df)
    outcome=pd.DataFrame(outcome)
    outcome.columns=["Transported"]
    outcome.index=origin_test_df["PassengerId"]
    outcome.replace(to_replace=1,value="True",inplace=True)
    outcome.replace(to_replace=0,value="False",inplace=True)
    outcome.to_csv(r'D:/desktop/spaceship_titanic/submission.csv',index=True
                   ,encoding="utf-8")

#用GBDT
def gbdt():
    gbc=GradientBoostingClassifier(random_state=1)
    gbc.fit(X_train, y_train)
    
    gbc.score(X_test,y_test)
    tmp1=pd.Series(gbc.feature_importances_,index=train_df.drop(['Transported'],axis=1).columns)
    tmp1=tmp1.sort_values(ascending=False)
    
    origin_test_df = pd.read_csv(r'D:/desktop/spaceship_titanic/test.csv')
    origin_test_df["PassengerId"]
    outcome=gbc.predict(test_df)
    outcome=pd.DataFrame(outcome)
    outcome.columns=["Transported"]
    outcome.index=origin_test_df["PassengerId"]
    outcome.replace(to_replace=1,value="True",inplace=True)
    outcome.replace(to_replace=0,value="False",inplace=True)
    outcome.to_csv(r'D:/desktop/spaceship_titanic/submission.csv',index=True
                   ,encoding="utf-8")

#确定要gbdt但是要确定超参
from sklearn.model_selection import GridSearchCV

param_test1 = {'n_estimators':range(20,200,10),
                'learning_rate':[0.1],
                "subsample":[1],
                "loss":["deviance"],
                "max_depth":[3],
                "min_samples_split":[2],
                "min_samples_leaf":[1],
                "min_weight_fraction_leaf":[0],
                "max_leaf_nodes":[None],
                "max_features":["sqrt",None]
               }

gs = GridSearchCV(estimator = GradientBoostingClassifier(random_state=10), 
                       param_grid = param_test1, scoring='roc_auc',cv=5)

X_train, y_train =(
    train_df.drop(['Transported'],axis=1), 
    train_df['Transported']
    )

gs.fit(X_train,y_train)
print(gs.best_params_)

######
origin_test_df = pd.read_csv(r'D:/desktop/spaceship_titanic/test.csv')
outcome=gs.predict(test_df)
outcome=pd.DataFrame(outcome)
outcome.columns=["Transported"]
outcome.index=origin_test_df["PassengerId"]
outcome.replace(to_replace=1,value="True",inplace=True)
outcome.replace(to_replace=0,value="False",inplace=True)
outcome.to_csv(r'D:/desktop/spaceship_titanic/submission.csv',index=True
               ,encoding="utf-8")















