# -*- coding: utf-8 -*-
"""
Created on Sun May 12 11:40:55 2019

@author: Administrator
"""
#import numpy as np
import pandas as pd
import os
import shutil
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

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
        outName=yearE+Qu_picName
       
        str1="\"<div>%s</div><img src=\"\"%s\"\" />\"" % (outName[:-4],outName)
        str2="\"<img src=\"\"%s\"\" />\"" % (yearE+An_picName)
        ankiData.append([str1,yearE,str2])
    
    



def copyfile(from_path,to_path,filedata,index=1):
    if index==1:
        from_path=os.path.join(from_path,'QU')
    else:
        from_path=os.path.join(from_path,'AN')
        
    if not os.path.exists(to_path):
        os.mkdir(to_path)
#print(os.path.join(from_path,filename1))
        
    for eachfile in filedata:
        shutil.copy(os.path.join(from_path,eachfile[0],eachfile[1]),os.path.join(to_path,eachfile[0]+eachfile[index]))
    #shutil.copy(os.path.join(from_path,fileName),os.path.join(to_path,asName))
    
def writeAnki_txt(to_filepath,data2write):
    with open(to_filepath, 'w') as f:
        for a in data2write:
            f.write(a[0]+"\t"+a[1]+"\t"+a[2]+"\t\n")#for improve


def writeDoc(Save_filepath,from_path,filedata):
    document = Document()                #以默认模板建立文档对象

    document.add_paragraph()
    p_detail = document.add_paragraph()
    r_detail = p_detail.add_run(u'真题错题集')
    r_detail.font.bold = True
    p_detail.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    from_path1=os.path.join(from_path,'QU')
    for eachfile in filedata:
        document.add_paragraph(eachfile[0]+eachfile[1].split('.')[0])
        document.add_picture(os.path.join(from_path1,eachfile[0],eachfile[1]),width=Inches(6))
        document.add_paragraph()
        document.add_paragraph()
        document.add_paragraph()
        

    p_detail = document.add_paragraph()
    r_detail = p_detail.add_run(u'答案')
    r_detail.font.bold = True
    p_detail.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    from_path2=os.path.join(from_path,'AN')
    for eachfile in filedata:
        document.add_paragraph(eachfile[0]+eachfile[2].split('.')[0])
        document.add_picture(os.path.join(from_path2,eachfile[0],eachfile[1]),width=Inches(4))
   
    document.save(Save_filepath)  

def copy_all_file(from_path,to_path,filedata):
    #lists = os.listdir(from_path)
    for eachfile in filedata:
        #print(i)
        shutil.copy(os.path.join(from_path,eachfile[0]+eachfile[1]),to_path)
        shutil.copy(os.path.join(from_path,eachfile[0]+eachfile[2]),to_path)  
    
def main():
    ankiData=[]
    ExcData=[]
    from_path="C:\\Users\\Administrator\\Documents\\Snagit"
    to_path=r"C:\Users\Administrator\AppData\Roaming\Anki2\000\collection.media"
    Exc_path="E:\\新建文件夹\\错题to_anki"
    Exc_filename="错题集WWX.xls"
    getExcData(os.path.join(Exc_path,Exc_filename),ExcData)
    
    formAnkiData(ExcData,ankiData)  
    print(ExcData)

    
    writeAnki_txt(os.path.join(Exc_path,"out.txt"),ankiData)
    copyfile(from_path,os.path.join(Exc_path,"media"),ExcData,1)
    copyfile(from_path,os.path.join(Exc_path,"media"),ExcData,2)
#    writeDoc(os.path.join(Exc_path,"错题.docx"),from_path,ExcData)
 
    copy_all_file(os.path.join(Exc_path,"media"),to_path,ExcData)    
    
    #print(ankiData)
 #   frame = pd.DataFrame(ankiData)
 #   frame.to_csv(os.path.join(to_path,"out.txt"))
    
main()
        
        
