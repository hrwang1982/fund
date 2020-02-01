import xlwt
import pandas as pd
import matplotlib.pyplot as plt
import time

# 这两行代码解决 plt 中文显示的问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

data = pd.read_excel('fund1.xls',sheet_name=0)
#data1=pd.DataFrame(data,columns=['序号','基金代码','基金简称','日期','自选','1年','2年','3年','从前'])
data1=pd.DataFrame(data)
df0 = data1.sort_values('自选',ascending=False).head(200)
df1m= data1.sort_values('最近1月',ascending=False).head(200)
df2_3m= data1.sort_values('最近2-3月',ascending=False).head(200)
df3m= data1.sort_values('最近3月',ascending=False).head(200)
df3_6m= data1.sort_values('最近4-6月',ascending=False).head(200)
df6m= data1.sort_values('最近6月',ascending=False).head(200)
df7_12m=data1.sort_values('最近7-12月',ascending=False).head(200)
df1y = data1.sort_values('最近1年',ascending=False).head(1500)
df1_2y = data1.sort_values('最近1-2年',ascending=False).head(1500)
df2_3y = data1.sort_values('最近2-3年',ascending=False).head(1500)
dfold = data1.sort_values('从前',ascending=False).head(1500)

# [最近一个月 & 最近2-3月] 的交集
df13m = pd.merge(df1m,df2_3m,how='inner')
# [ 最近1个月 & 最近2-3月 & 最近4-6月] 的交集
df136m = pd.merge(df13m,df3_6m,how='inner')
# [最近1个月 & 最近2-3月 & 最近4-6月 & 最近7-12月] 的交集
dfm= pd.merge(df136m,df7_12m,how='inner')

# [最近1年 & 最近1-2年] 的交集
df12y = pd.merge(df1y,df1_2y,how='inner')
# [最近1年 & 最近1-2年 & 最近2-3年] & 的交集
df123y = pd.merge(df12y,df2_3y,how='inner')
# [最近1年 & 最近1-2年 & 最近2-3年 & 从前] & 的交集
dfy = pd.merge(df123y,dfold,how='inner')

dff = pd.merge(dfm,dfy,how='inner')

dfm.to_excel('fund3.xls',sheet_name='fund3',encoding='utf-8')

dfy.to_excel('fund4.xls',sheet_name='fund4',encoding='utf-8')

pdata=pd.DataFrame(dfm,columns=['序号','基金代码','基金简称','日期','自选','最近1月','最近2-3月','最近4-6月','最近7-12月','最近1-2年','最近2-3年'])
xtime = pdata.columns.values[5:]
for index,row in pdata.iterrows():
    #print(list(row)[5:-1])
    plt.plot(xtime,list(row)[5:], '.-', label=list(row)[1:3])
plt.xticks(xtime)
plt.xlabel('周期')
plt.ylabel('涨幅百分比')
plt.legend()
plt.savefig( time.strftime("%Y-%m-%d", time.localtime()) +".png" )
plt.show()
