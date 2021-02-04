        # -*- coding: utf-8 -*-
import sys
import re
from tkinter import *
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import pandas as pd
import numpy as np

def fuzzy_merge(df_1, df_2, key1, key2, threshold,self,limit=1):

    self.lb3.insert(END,"檔案處理中...")
    
    df_1[key1] = df_1[key1].fillna(method='ffill')
    df_2[key2] = df_2[key2].fillna(method='ffill')
    df_1 = df_1.drop_duplicates(subset = key1, keep = "first")
    df_2 = df_2.drop_duplicates(subset = key2, keep = "first")
    
    s = df_2[key2].tolist()
    
    m = df_1[key1].apply(lambda x: process.extract(x, s, limit=limit))    
    df_1['matches'] = m

    self.progress['value'] = 20
    
    num = df_1['matches'].apply(lambda x: ', '.join([str(i[1]) for i in x if i[1] >= threshold]))
    df_1['similarity'] = num

    self.progress['value'] = 40

    m2 = df_1['matches'].apply(lambda x: ', '.join([i[0] for i in x if (i[1] >= threshold)]))
    df_1['matches'] = m2

    self.progress['value'] = 60
    
    df_1['numKey1'] = df_1[key1].apply(lambda x: re.findall("[一二三五六七八九十]",x))
    df_1['numKey2'] = df_1['matches'].apply(lambda x: re.findall("[一二三五六七八九十]",x))
    for i in range(0,len(df_1['numKey1'])):
        if df_1['numKey1'][i] != df_1['numKey2'][i]:
            df_1.loc[i, 'matches'] = '' 
            df_1.loc[i, 'similarity'] = ''
            
    self.progress['value'] = 80
    
    if 'numKey1' in df_1.columns:
        df_1 = df_1.drop('numKey1', 1)
    if 'numKey2' in df_1.columns:
        df_1 = df_1.drop('numKey2', 1)
    
    return df_1

def exceldealfunc(filename1,filename2,filename3,num1,num2,num3,self):
    f1  = pd.read_excel(filename1)
    f2  = pd.read_excel(filename2)
    self.lb3.insert(END,"檔案讀取完成")
    self.progress['value'] = 10
    if num3 == 100:
        merged = f1.merge(f2, left_on = f1.columns[num1], right_on = f2.columns[num2], how = 'left')
    else:
        f1Merged = fuzzy_merge(f1, f2, f1.columns[num1], f2.columns[num2],num3,self)
        merged = f1Merged.merge(f2, left_on = 'matches', right_on = f2.columns[num2], how = 'left')
    with pd.ExcelWriter(filename3) as writer:
        merged.to_excel(writer, sheet_name='Result')
    self.lb3.insert(END,"檔案寫入完成")
    self.progress['value'] = 100

