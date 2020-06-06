import xlwt
import pandas as pd
import matplotlib.pyplot as plt
import time

my_funds=("005275","162605","001076","110011","270050","000083","519674","486001")
# 这两行代码解决 plt 中文显示的问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

data = pd.read_excel('fund1.xls',sheet_name=0)
#data1=pd.DataFrame(data,columns=['序号','基金代码','基金简称','日期','自选','1年','2年','3年','从前'])
data1=pd.DataFrame(data)


pdata=pd.DataFrame(data1,columns=['序号','基金代码','基金简称','日期','自选','最近1月','最近2-3月','最近4-6月','最近7-12月','最近1-2年','最近2-3年'])
pdata1=pd.DataFrame(columns=['序号','基金代码','基金简称','日期','自选','最近1月','最近2-3月','最近4-6月','最近7-12月','最近1-2年','最近2-3年'])
xtime = pdata.columns.values[5:]
for fund in my_funds:
    print(fund)
    for index, row in pdata.iterrows():
        #print(list(row)[5:-1])
        if fund == str(row[1]):
            print(row[1],row[2])
            row1=dict(row)
            pdata1=pdata1.append(row1,ignore_index=True)
            plt.plot(xtime,list(row)[5:], '.-', label=list(row)[1:3])
plt.xticks(xtime)
plt.xlabel('周期')
plt.ylabel('涨幅百分比')
plt.legend()
plt.savefig( time.strftime("%Y-%m-%d", time.localtime()) +".png" )
plt.show()
print(pdata1)