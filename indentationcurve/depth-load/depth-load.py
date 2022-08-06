import xlwt
import os
import pandas as pd
import numpy as np
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

def txt_xls(filename,xlsname):
    """
    :文本转换成xls的函数
    :param filename txt文本文件名称、
    :param xlsname 表示转换后的excel文件名
    """
    try:
        f = open(filename) 
        xls=xlwt.Workbook()
        #生成excel的方法，声明excel
        sheet = xls.add_sheet('sheet1',cell_overwrite_ok=True) 
        x = 0 
        while True:
            #按行循环，读取文本文件
            line = f.readline() 
            if not line: 
                break  #如果没有内容，则退出循环
            for i in range(len(line.split('\t'))):
                item=line.split('\t')[i]
                sheet.write(x,i,item) #x单元格经度，i 单元格纬度
            x += 1 #excel另起一行
        f.close()
        xls.save(xlsname) #保存xls文件
    except:
        raise

def datacleaning(filelist,startrow,depthcolumn,loadcolumn):

    #数据清洗
    #输入待处理的文件列表
    #得到处理之后的汇总的excel
    #返回文件名组成的列表作为画图时的图例使用
    '''
    需要指明数据开始的行，以及深度和载荷的列
    一般同一组试验仪器得到的数据都是一致的
    数据开始的行为整数，数据开始的列为字符串
    '''
    
    frameout=pd.DataFrame()
    labelslist=[]
    for filepath in filelist:
        frame=pd.read_excel(filepath,header=None,skiprows=startrow-1,usecols=depthcolumn+','+loadcolumn)
        series=frame[0]
        series=series[series==0]
        rownumber0=series.index[0] #深度等于0的行
        frame=frame.iloc[rownumber0:,:]
        frame=frame[(frame[1]>0)&(frame[0]>=0)]#确保深度和载荷大于零
        filename=os.path.basename(filepath)
        labelslist.append(filename)
        A='depth'+filename
        B='load'+filename
        frame.columns=[A,B]
        frame.index=frame.index-rownumber0#从零开始
        frameout=pd.concat([frameout,frame],axis=1)
        # frameout=frameout.fillna(method='pad',axis=0)可以不用处理NaN
    return frameout,labelslist

def pplot(frameout,labelslist):
#绘图函数

    fig=plt.figure(num=1, figsize=(15, 7))
    font=fm.FontProperties(fname = r'C:\Windows\Fonts\times.ttf',size = 15)

    xm=2000
    ym=10
    xd=100
    yd=1
    plt.xlim(0,xm)
    plt.ylim(0,ym)
    x_ticks=np.arange(0,xm,xd)
    y_ticks=np.arange(0,ym,yd)
    plt.xticks(x_ticks,fontproperties=font)
    plt.yticks(y_ticks,fontproperties=font)
    plt.xlabel('depth(nm)',fontproperties=font)
    plt.ylabel('load(mN)',fontproperties=font)

    #对不同载荷下
    # plt.axhline(0.025,c='r',alpha = 0.5)
    # plt.axhline(0.5,c='r',alpha = 0.5)
    # plt.axhline(1,c='r',alpha = 0.5)
    # plt.axhline(5,c='r',alpha = 0.5)
    # plt.axhline(10,c='r',alpha = 0.5)
    # plt.axhline(30,c='r',alpha = 0.5)
    # plt.axhline(50,c='r',alpha = 0.5)
    # plt.axhline(100,c='r',alpha = 0.5)
    # plt.axhline(200,c='r',alpha = 0.5)


    for i in range(len(labelslist)):
        x=frameout.iloc[:,2*i]
        y=frameout.iloc[:,2*i+1]/1000
        plt.plot(x,y,label=labelslist[i])

    plt.legend()

    plt.axhline(10,c='r',alpha = 0.5)
    fig.savefig(path+'1.svg', dpi=600, pil_kwargs={'quality': 95})
    plt.show()
    
path=input('请输入txt文件夹的完整路径：')
filelist=getpathlist(path,'txt')
for filename in filelist:
    xlsname  = filename[:-4]+".xls"
    txt_xls(filename,xlsname)

filelist=getpathlist(path,'xls')
frameout,labelslist=datacleaning(filelist,7,'A','B')
frameout.to_excel(path + "\\results.xlsx")
pplot(frameout,labelslist)

