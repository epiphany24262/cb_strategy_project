# -*- coding: utf-8 -*-
"""
Created on Sun Mar 29 10:06:11 2026

@author: dhhly
"""

import glob
import numpy as np
import pandas as pd
import sys
import statsmodels 
import statsmodels.api
import matplotlib.pyplot as plt
import datetime

import warnings

#step1 读取数据
print('lude data start!!!!!!!!!!!!!!!')
df = pd.read_parquet('cb_data.pq')  
df = df.reset_index()
df['trade_date'] = pd.to_datetime(df['trade_date'])  #交易日期是字符串，需要转化为datetime的格式


#一些简单的预处理
code_to_name = dict(zip(df["code"], df["name"]))   #这个code to name的dict方便未来


#简单的回测框架

#2 构造价格矩阵
pricedf = df[['trade_date','code','close']]  
pricedf = pricedf.set_index(['trade_date','code']).unstack()['close'] #unstck 的作用是进行行列转换!!!

#3接下来用价格矩阵来构造收益率矩阵，直接用DataFrame的pct_change函数就可以了。如果因子里面有，那么可以直接用
day_return = pricedf.pct_change().shift(-1)
day_return2 = df[['trade_date','code','pct_chg']].set_index(['trade_date','code']).unstack()['pct_chg'].shift(-1)  


#3构造因子矩阵,这里以溢价率为例子conv_prem
conv_prem_df = df[['trade_date','code','conv_prem']]  
conv_prem_df = conv_prem_df.set_index(['trade_date','code']).unstack()['conv_prem'] 


#3.1 分组测试
def group_analysis(factor,forward_return,num=5,title=""):
    cutfactor=factor.rank(axis=1).apply(lambda x:pd.qcut(x,num,range(1,num+1)),axis=1)    #分组标签
    result=pd.DataFrame()    
    for group in range(1,1+num):        
        group_return=forward_return[cutfactor==group].sum(axis=1)/(cutfactor==group).sum(axis=1)        
        result[group]=group_return    
    result['base']=(forward_return[~factor.isna()]).sum(axis=1)/(~factor.isna()).sum(axis=1)    
    (result+1).cumprod().plot(figsize=(12,5),grid=True,title=title);
    return result
group_analysis(conv_prem_df,day_return,title="conv_prem");


factor=conv_prem_df+ pricedf  #合成信号因子。但是这里需要指出，这样其实很麻烦，每个因子都需要单独df来处理。


#4 根据排名进行筛选

N=20

def selectTopN(tmp):  #这个函数作用是选出这一行最小的N个数字，
                      #并且把他们的index记下后，赋值为1，其他都为0
    tmp=tmp.copy()
    symbols=tmp.nsmallest(N).index    
    tmp[:]=0    
    tmp[symbols]=1    
    return tmp

signal=factor.apply(selectTopN,axis=1)  #这个apply函数加上axis=1是将函数作用于input的每一行

#5 计算每一天的收益 并 画出收益曲线
pnl=(signal*day_return).sum(axis=1)/N
(pnl+1).cumprod().plot(figsize=(10,5),grid=True)
















