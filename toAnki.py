# -*- coding: utf-8 -*-
"""
Created on Sun May 12 11:40:55 2019

@author: Administrator
"""
#import numpy as np
import pandas as pd
import os
import shutil

def getExcData(Exc_path,ExcData):
    df1 = pd.read_excel(Exc_path)
    df2 = df1.dropna(thresh=2)
    for index in df2.index:
        yearErr=df2.loc[index].dropna().values
        yearE=yearErr[0]
        yearQnum=yearErr[1:]
    #print(yearE)
        for Enum in yearQnum:
            QpicName="%02d.png" % Enum
            ApicName="%02dAN.png" % Enum
            ExcData.append([yearE,QpicName,ApicName])
    
    
    
def formAnkiData(fromData,ankiData):
    
    for Enum in fromData:
        yearE= Enum[0]
        Qu_picName= Enum[1]
        An_picName= Enum[2]
        
       
        str1="\"<img src=\"%s\" />" % (yearE+Qu_picName)
        str2="\"<img src=\"%s\" />\"" % (yearE+An_picName)
        ankiData.append([str1,yearE,str2])
    
    



def copyfile(from_path,fileName,to_path,asName):
    if not os.path.exists(to_path):
        os.mkdir(to_path)
#print(os.path.join(from_path,filename1))
    shutil.copy(os.path.join(from_path,fileName),os.path.join(to_path,asName))
    
    
    
def main():
    ankiData=[]
    ExcData=[]
    from_path="C:\\Users\\Administrator\\Documents\\Snagit\\2004S"
    to_path="D:\\media"
    Exc_path="E:\\新建文件夹\\错题to_anki\\错题20190512.xlsm"
    getExcData(Exc_path,ExcData)
    formAnkiData(ExcData,ankiData)
    print(ExcData)
    print(ankiData)
    frame = pd.DataFrame(ankiData)
    frame.to_csv(os.path.join(to_path,"out.txt"))
    
main()
        
        
