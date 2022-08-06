import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from scipy.optimize import curve_fit


def getpathlistofexcel(path):
    path_list=os.listdir(path)
    for filename in path_list:
        path_list[path_list.index(filename)] = os.path.join(path,filename)
    path_listexcel = []
    for filename in path_list:
        if filename[-4:] == 'xlsx':
            path_listexcel.append(filename)
    path_listnumber = len(path_listexcel)
    return path_listexcel,path_listnumber

def average(data):
    for i in range(2):
        frame1 = pd.DataFrame()
        frame2 = pd.DataFrame()
        for j in range(5):
            columnnumber = j+5*i
            x = data.iloc[:,2*columnnumber+1]
            y = data.iloc[:,2*columnnumber]
            frame1 = pd.concat([frame1,x],axis = 1)
            frame2 = pd.concat([frame2,y],axis = 1)
        lis1 = []
        for index in frame1.index:
            series = frame1.iloc[index,:]
            ave = series.mean()
            lis1.append(ave)
        frame1['average'] = lis1
        lis2 = []
        for index in frame2.index:
            series = frame2.iloc[index,:]
            ave = series.mean()
            lis2.append(ave)
        frame2['average'] = lis2
        frame1.to_excel(filepath + "\\"+str(i)+"time.xlsx")
        frame2.to_excel(filepath + "\\"+str(i)+"depth.xlsx")

def mainfunc(filelist):
    fig = plt.figure(num=1, figsize=(14, 7))
    font = fm.FontProperties(fname = r'C:\Windows\Fonts\times.ttf')

    xm = 110
    xd = 10
    ym = 160
    yd = 20
    plt.xlim(0,xm)
    plt.ylim(0,ym)
    x_ticks = np.arange(0,xm,xd)
    y_ticks = np.arange(0,ym,yd)
    plt.xticks(x_ticks,fontproperties = font,size = 20)
    plt.yticks(y_ticks,fontproperties = font,size = 20)
    plt.xlabel('time(s)',fontproperties = font,fontsize = 25)
    plt.ylabel('depth(nm)',fontproperties = font,fontsize =25)

    for i in range(2):
        frame1 = pd.read_excel(filelist[2*i+1])
        frame2 = pd.read_excel(filelist[2*i])
        y = np.asarray(frame2['average'])
        y = y-y[0]
        y = y[0::100]
        x = np.asarray(frame1['average'])
        x = x-x[0]
        x = x[0::100]

        def func(x,a,b,c):
            return a*(x**b)+c*x
        popt, pcov = curve_fit(func, x, y)
        #print(type(popt))
        #popt是numpy.ndarray类型的数组
        
        if i == 0:
            colors = 'r'
            number = 'L1'
            delta = 0
        else:
            colors = 'b'
            number = 'M1'
            delta = 0

        plt.scatter(x, y,marker='.',c='none',edgecolors='black')

        mean = np.mean(y)  # 1.y mean
        ss_tot = np.sum((y - mean) ** 2)  # 2.total sum of squares
        ss_res = np.sum((y - func(x, *popt)) ** 2)  # 3.residual sum of squares
        r_squared = 1 - (ss_res / ss_tot)  # 4.r squared
        plt.plot(x,func(x, *popt),c = colors,label = number +' '
        'a=' + str('{:.3f}'.format(popt[0])) +
        ',m=' + str('{:.3f}'.format(popt[1])) + 
        ',k=' + str('{:.3f}'.format(popt[2])) + ' '
        '$\mathregular{R^2}$=' + str('{:.4f}'.format(r_squared-delta)))
        plt.legend(fontsize = 15)
    fig.savefig(filepath+'1.svg', dpi=600)

filepath = input('请输入文件夹的路径：')
data = pd.read_excel(filepath + "\\results.xlsx").iloc[:,1:]
average(data)
filelist,m= getpathlistofexcel(filepath)
mainfunc(filelist)
print('Done!')