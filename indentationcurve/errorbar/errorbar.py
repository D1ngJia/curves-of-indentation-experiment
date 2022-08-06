import numpy as np
import pandas as pd
import os
import xlwt
import math
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

def getpathlist(path,suffix):
#读取一个文件夹下所有.xls或者.txt文件
#path为文件夹路径，suffix为需要读取的文件类型后缀
#返回所有该类型文件的绝对路径组成的列表

    path_list=os.listdir(path)
    path_listout=[]
    for filename in path_list:
        if filename[-len(suffix):]==suffix:
            path_listout.append(os.path.join(path,filename))
    return path_listout

def datacleaning(excellist,depthcolunm,colunm_E,colunm_H,rownumber,startrow):#输入文件夹路径
    #数据清洗
    #输入待处理的文件列表
    #得到处理之后的汇总的深度，模量和硬度的三个excel
    '''
    需要指明数据开始的行，以及深度，模量和硬度的列
    一般同一组试验仪器得到的数据都是一致的
    数据开始的行和三个数据列均为从1开始的整数
    rownumber为有效数据的总行数
    '''
    frameout_x=pd.DataFrame()
    frameout_H=pd.DataFrame()
    frameout_E=pd.DataFrame()
    columnslist=[]
    for file in excellist:
        data_x=pd.read_excel(file,skiprows=startrow-1,header=None).iloc[:,depthcolunm-1]
        data_E=pd.read_excel(file,skiprows=startrow-1,header=None).iloc[:,colunm_E-1]
        data_H=pd.read_excel(file,skiprows=startrow-1,header=None).iloc[:,colunm_H-1]
        frameout_x=pd.concat([frameout_x,data_x],axis=1)
        frameout_E=pd.concat([frameout_E,data_E],axis=1)
        frameout_H=pd.concat([frameout_H,data_H],axis=1)
        filename=os.path.basename(file)
        columnslist.append(filename)
    frameout_x.columns=columnslist
    frameout_E.columns=columnslist
    frameout_H.columns=columnslist
    ls_x=[]
    for i in range(rownumber):
        series=frameout_x.iloc[i,:]
        ave_x=series.mean()
        ls_x.append(ave_x)
        #复习增加一列
        frameout_x['average']=pd.Series(ls_x)
    ls_ave_H=[]
    ls_se_H=[]
    ls_ave_E=[]
    ls_se_E=[]
    for i in range(rownumber):
        series_H=frameout_H.iloc[i,:]
        ave_H=series_H.mean()
        #计算标准误
        se_H=np.std(series_H)/math.sqrt(len(columnslist))
        ls_ave_H.append(ave_H)
        ls_se_H.append(se_H)
        series_E=frameout_E.iloc[i,:]
        ave_E = series_E.mean()
        se_E = np.std(series_E)/math.sqrt(len(columnslist))
        ls_ave_E.append(ave_E)
        ls_se_E.append(se_E)

        frameout_H['average_h']=pd.Series(ls_ave_H)
        frameout_H['se_h']=pd.Series(ls_se_H)
        frameout_E['average_e']=pd.Series(ls_ave_E)
        frameout_E['se_e']=pd.Series(ls_se_E)
    return frameout_x,frameout_H,frameout_E

def pplot():
    #在相同的路径下绘制误差棒曲线
    frame1=pd.read_excel(filepath + "\\resultsx.xlsx")
    frame2=pd.read_excel(filepath + "\\resultsH.xlsx")
    frame3=pd.read_excel(filepath + "\\resultsE.xlsx")

    x=frame1['average']
    h=frame2['average_h']
    dh=frame2['se_h']
    e=frame3['average_e']
    de=frame3['se_e']

    fig=plt.figure(num=1, figsize=(14, 7))
    font=fm.FontProperties(fname = r'C:\Windows\Fonts\times.ttf')

    xm = 1400
    ym = 2.2
    xd = 100
    yd = 0.2
    plt.xlim(0,xm)
    plt.ylim(0,ym)
    x_ticks = np.arange(0,xm,xd)
    y_ticks = np.arange(0,ym,yd)
    plt.xticks(x_ticks,fontproperties=font,fontsize=20)
    plt.yticks(y_ticks,fontproperties=font,fontsize=20)
    plt.xlabel('depth(nm)',fontproperties=font,fontsize=25)
    plt.ylabel('hardness(GPa)',fontproperties=font,fontsize=25)
    plt.errorbar(x,h,yerr=dh,fmt='.b',ecolor='gray',elinewidth=1,capsize=1)
    fig.savefig(filepath+'\\H.jpg', dpi=600, pil_kwargs={'quality': 95})

    plt.clf()

    xm = 1400
    ym = 16
    xd = 100
    yd = 2
    plt.xlim(0,xm)
    plt.ylim(0,ym)
    x_ticks = np.arange(0,xm,xd)
    y_ticks = np.arange(0,ym,yd)
    plt.xticks(x_ticks,fontproperties=font,fontsize=20)
    plt.yticks(y_ticks,fontproperties=font,fontsize=20)
    plt.xlabel('depth(nm)',fontproperties=font,fontsize=25)
    plt.ylabel('modulus(GPa)',fontproperties=font,fontsize=25)
    plt.errorbar(x,e,yerr=de,fmt='.r',ecolor='gray',elinewidth=1,capsize=1)
    fig.savefig(filepath+'\\E.jpg', dpi=600, pil_kwargs={'quality': 95})
    print('Done!')

filepath=input('请输入excel文件夹的完整路径：')

excellist=getpathlist(filepath,'xls')
frame1,frame2,frame3=datacleaning(excellist,3,35,33,61,4)
frame1.to_excel(filepath + "\\resultsx.xlsx")
frame2.to_excel(filepath + "\\resultsH.xlsx")
frame3.to_excel(filepath + "\\resultsE.xlsx")
pplot()





