#!/usr/bin/env python
# coding: utf-8

# In[120]:


# encoding: UTF-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlwt
import xlrd
from xlutils.copy import copy
import pandas as pd
import matplotlib.pyplot as plt
import time
import pandas as pd
import datetime as datetime
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
import threading
import openpyxl


#使用的chrome驱动器,同一个driver复制多份儿，用于多线程
chrome_driver=r'C:\工作\pychram\111\chromedriver.exe'
chrome_driver1=r'C:\工作\pychram\111\chromedriver1.exe'
chrome_driver2=r'C:\工作\pychram\111\chromedriver2.exe'

# 这两行代码解决 plt 中文显示的问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

#调用多线程抓取交易明细函数getfund_mingxi_mt ， 需要使用全局变量，如下两个变量为此定义
tmp_funds1=pd.DataFrame(columns=["基金代码","基金简称","净值日期", "单位净值", "累计净值", "日增长率"])

#########################################################
### **** 这个每次需要更改一下加载进来的全量文件 **** ###
#########################################################
#将上次的读取的全量基金详情加载进来
tmp_funds1=pd.read_excel(r"C:\fund_mx_get\all_fund_20200714_1.xlsx", encoding='utf-8')
#将包含全量基金明细数据中的基金代码都转换成6位，左侧补0，如5275补充完后是005275
tmp_funds1['基金代码']= tmp_funds1['基金代码'].map(lambda x : str(x).zfill(6) ) 

#这个k是多线程抓取所有基金信息的初始index，因为要将新抓的合并程一个统一的文件，所以先读取老的，然后基于老的继续添加
k=len(tmp_funds1)

#定义多线程的锁
lock = threading.Lock()

#打开chrome浏览器，打开指定的url，并返回driver
def open_url(url):
    #options = webdriver.ChromeOptions()
    #options.add_argument('--no-sandbox')
    #options.add_experimental_option('excludeSwitches', ['enable-automation'])
    #driver = webdriver.Chrome(executable_path=chrome_driver, options=options)

    driver = webdriver.Chrome(executable_path=chrome_driver)
    driver.get(url)
    driver.implicitly_wait(20)
    return driver


# In[36]:


'''
如果基金的历史数据保存在多个文件中，可以将文件内容读取出来，合并成一个pd，这样便于对长期的历史数据进行分析
'''
file_list=["0","1","2","3","4","5","6"]
folder_dir= "C:/fund_mx_get"
def read_file(file_list):
    df_list= []
    for cur_file in file_list:
        print(cur_file)
        source_file='{}/2020-07-05_funds_mingxi{}.xls'.format(folder_dir,cur_file)
        df0 = pd.read_excel(source_file, header=0, sep=',', encoding='utf-8')
        df_list.append(df0)
    return df_list
df_list = read_file(file_list)
df0=pd.concat(df_list)
df0=df0[['基金代码','基金简称','净值日期','单位净值','累计净值','日增长率']]
before_all_fund=df0.drop_duplicates()


# In[37]:


def late_time(time2,ndays):
    #time2是外部传入的任意日期
    now_time = datetime.datetime.strptime(time2, '%Y-%m-%d')
    #如需求是当前时间则去掉函数参数改写      为datetime.datetime.now()
    threeDayAgo = (now_time - datetime.timedelta(days = ndays))
    # 转换为时间戳
    timeStamp =int(time.mktime(threeDayAgo.timetuple()))
    # 转换为其他字符串格式
    otherStyleTime = threeDayAgo.strftime("%Y-%m-%d")
    return otherStyleTime

month1 = late_time("2020-6-19",30)
month2 = late_time("2020-6-19",60)
month3 = late_time("2020-6-19",90)
month4 = late_time("2020-6-19",120)
month5 = late_time("2020-6-19",150)
month6 = late_time("2020-6-19",180)


# In[121]:


'''
传入一个包含全部基金的基金代码和基金简称的pd。 再传进去一个包含自己持有基金的基金代码的列表.
返回自己持有基金的基金代码和基金简称的pd。
'''
def fund_num_name(alllist_pd,*myfunds_num_list):
    r_num_name=pd.DataFrame(columns=['基金代码','基金简称'])
    if len(myfunds_num_list) > 0:
        i=0
        for mf in myfunds_num_list:
            mf=int(mf)      
            list = alllist_pd['基金简称'][alllist_pd['基金代码'] == mf].tolist()
            list.insert(0,mf)                     
            r_num_name.loc[i]=list
            i=i+1
    else:
        r_num_name=alllist_pd[['基金代码','基金简称']]
    
    return r_num_name


# In[122]:


#无线程处理， 传入包含要抓取基金的pd，然后爬取最近160个工作日详细的每日净值信息
def getfund_mingxi(funds, filename):

    tmp_funds=pd.DataFrame(columns=["基金代码","基金简称","净值日期", "单位净值", "累计净值", "日增长率"])
    #索引开始数字
    k=0

    #接下来execl将从该行开始写入
    m = 1

    for index,row in funds.iterrows():
        
        fund=str(row['基金代码']).zfill(6)
        print(fund)
        name=row['基金简称']
        url="http://fundf10.eastmoney.com/jjjz_"+ fund + ".html"
        print("******url*******")
        print(url)
        driver.get(url)
        driver.implicitly_wait(20)

        #获取基金最新的日期
        lasttran_date= driver.find_element_by_xpath('//*[@id="jztable"]/table/tbody/tr[1]/td[1]').text
        print(lasttran_date)

        #从第一页开始抓取，每页包含20个工作日的净值信息，抓取8页的详细信息，包含了最近半年的详情
        y=1 
        while y < 9:
            if y > 1:
                try:
                    print("******click*******")
                    driver.find_element_by_xpath('//*[@id="pagebar"]/div[1]/label[8]').click()  
                    time.sleep(3)
                except NoSuchElementException:
                    driver.find_element_by_xpath('//*[@id="pagebar"]//*[text()="下一页"]').click()
                    print("Choose another way to click next page")
                except WebDriverException:
                    driver.refresh()
                    print("Refresh this page")
                else:
                    pass
                
            table_context = driver.find_element_by_xpath('//*[@id="jztable"]/table/tbody').find_elements_by_tag_name("tr")

            #将抓取的信息存放到pd中

            for line in table_context:
                fund_list=[fund,name]
                clos = line.find_elements_by_tag_name("td")
                n = 1
                for item in clos[:4]:
                    text = item.text
                    fund_list.append(text)

                tmp_funds.loc[k]=fund_list
                k = k + 1
            y += 1

    #将基金交易明细数据保存到execl
    tmp_funds.to_excel(time.strftime("%Y%m%d", time.localtime()) + "_" + filename + ".xlsx", index=False)
    
    return tmp_funds


# In[123]:


#mt-multi thread多线程支持，传入包含要抓取基金的pd，然后多线程爬取最近n页的净值信息(每页包含20个工作日的净值信息)
def getfund_mingxi_mt(funds,driver,pagenum):
    #tmp_funds1=pd.DataFrame(columns=["基金代码","基金简称","净值日期", "单位净值", "累计净值", "日增长率"])
    #多线程需要将处理的数据合并，需要用到全局变量
    global tmp_funds1
    #索引开始数字
    global k

    #抓取最近N页的净值信息， 如最近2页，因为有for循环 <n 的判断，<2 则只抓取1页，所以需要n+1 
    pagenum = pagenum + 1
    for index,row in funds.iterrows(): 
        fund=str(row['基金代码']).zfill(6)
        print(fund)
        name=row['基金简称']
        url="http://fundf10.eastmoney.com/jjjz_"+ fund + ".html"
        print(name,url)

        driver.get(url)
        driver.implicitly_wait(20)

        #获取基金最新的日期
        lasttran_date= driver.find_element_by_xpath('//*[@id="jztable"]/table/tbody/tr[1]/td[1]').text
        print(lasttran_date)

        #从第一页开始抓取，每页包含20个工作日的净值信息，抓取n页的详细信息，包含了最近半年的详情
        y=1 
        while y < pagenum:
            if y > 1:
                try:
                    print(name,"******click*******")
                    driver.find_element_by_xpath('//*[@id="pagebar"]/div[1]/label[pagenum]').click()  
                    time.sleep(3)
                except NoSuchElementException:
                    driver.find_element_by_xpath('//*[@id="pagebar"]//*[text()="下一页"]').click()
                    print("Choose another way to click next page")
                except WebDriverException:
                    driver.refresh()
                    print("Refresh this page")
                else:
                    pass
                
            table_context = driver.find_element_by_xpath('//*[@id="jztable"]/table/tbody').find_elements_by_tag_name("tr")
            print(name,"linenum",len(table_context))

            #将抓取的信息存放到pd中
            for line in table_context:
                fund_list=[fund,name]
                clos = line.find_elements_by_tag_name("td")
                n = 1
                for item in clos[:4]:
                    text = item.text
                    fund_list.append(text)

                #tmp_funds.loc[k]=fund_list
                print(name,fund_list)
                lock.acquire()
                try:
                    #tmp_funds1=tmp_funds1.append(fund_list)
                    tmp_funds1.loc[k]=fund_list
                    k = k + 1
                finally:
                    lock.release()
                
            y += 1
    return tmp_funds1


# In[124]:


'''
对传入的fundlist pd中的基金进行净值涨幅计算,该fundlist中的基金数据是串行的，时间倒序的，type分为按month、week进行计算，daytype指工作日WD还是自然日CD，daynum指month或者week中间隔的日子数量
该函数计算最近26个周的每周涨幅， 或者计算最近6个月的每月涨幅
'''
def fund_rate(fundlist,type,daytype,daynum):
    #将fundlist中基金代码去重，并转换成列表
    list1=fundlist['基金代码'].drop_duplicates().values.tolist()
    global lasttrans_day

    #按照工作日来处理，那么是根据fundlist数据集中的index来定位。 处理daynum=5 则对应着week，daynum=22则对应着month，获取对应数据
    if daytype=="WD":
        #创建一个pd，用于存放计算出来的每月涨幅数据
        fund_xx=fundlist[fundlist['基金代码']=="005275"]
        fund_idxname=[]
        fund_index1 = fund_xx.index[0]
        fund_day = fund_xx.loc[fund_index1,'净值日期']    
        fund_idxname.append(fund_day)
        #计算周数据，查看最近26周的数据
        if type=="week":
            howlong=26
        #否则月数据，查看最近6个月的数据
        if type=="month":
            howlong=6
            
        m = 0
        while m < howlong :
                fund_index1 =  fund_index1 + daynum
                fund_day = fund_xx.loc[fund_index1,'净值日期']
                fund_idxname.append(fund_day)
                m = m + 1

        fund_idxname.pop()
        fund_idxname.insert(0,'最新日期')
        fund_idxname.insert(0,'基金简称')
        fund_idxname.insert(0,'基金代码')
        fundrate_result = pd.DataFrame(columns=fund_idxname)

        #保存基金涨幅的索引号
        j = 0 
        #开始计算每个基金的涨幅
        for fund in list1:
            fund_xx=fundlist[fundlist['基金代码']==fund]
            fund_indexs = []
            fund_rates = []

            m=0
            fund_index1 = fund_xx.index[0]
            fund_indexs.append(fund_index1)
            while m < howlong :
                fund_index1 =  fund_index1 + daynum
                fund_indexs.append(fund_index1)
                first_value = fund_xx.loc[fund_indexs[m],'单位净值']
                second_value = fund_xx.loc[fund_indexs[m+1],'单位净值']
                first_value = float(first_value)
                second_value = float(second_value)
                first_rate =  round((first_value - second_value)/second_value * 100,2)
                fund_rates.append(first_rate)
                m = m + 1
            name = fund_xx.loc[fund_indexs[m],'基金简称']
            fund_rates.insert(0,lasttrans_day)
            fund_rates.insert(0,name)
            fund_rates.insert(0,fund)
            fundrate_result.loc[j]=fund_rates
            j = j + 1

    #按照datetype=CD来处理
    else:
        print("I don't finish this part")
        #if type=="week":
        
        #按照type=="month"来处理
        #else:
     
    return fundrate_result


# In[125]:


'''
多线程抓取基金明细后合并的结果fundlist pd，对该pd中的基金进净值涨幅计算,该fundlist中的基金数据是非串行的，时间倒序的。
所以不能用索引名字（默认行号）进行基金净值获取，如当前index的名称100，不是加5，105就是上周基金净值所在行。
过滤出某只基金的明细数据，iloc + 5 一定是前一周的。
type分为按month、week进行计算; daytype指工作日WD还是自然日CD，daynum指month或者week中间隔的日子数量,period_n 只计算出多少个周期的
'''
def fund_rate_mt(fundlist,type,daytype,daynum,period_n):
    #将fundlist中基金代码去重，并转换成列表
    list1=fundlist['基金代码'].drop_duplicates().values.tolist()
    global lasttrans_day

    #按照工作日来处理，那么是根据fundlist数据集中的index来定位。 处理daynum=5 则对应着week，daynum=22则对应着month，获取对应数据
    if daytype=="WD":
        #创建一个pd，用于存放计算出来的每月涨幅数据
        fund_xx=fundlist[fundlist['基金代码']=="005275"]
        fund_idxname=[]
        
        #计算周数据，查看最近period_n周的数据，一般短期看最近6周的就够了。 同时要计算一下不同类型需要的最少数据是多少
        if type=="week":
            howlong=period_n
            min_count = daynum * period_n + 2 
        #否则月数据，查看最近period_n个月的数据，一般中长期看最近6个月的就够了
        if type=="month":
            howlong=period_n
            min_count = daynum * period_n + 2 
            
        #单独过滤出来的某只基金的pd，索引编号都是0开始，索引名称并不一定是顺序的（因为多线程抓取的原因）,所以下面都用iloc，
        #列从0开始，2表示第3列，是净值日期
        fund_index1 = 0  
        fund_day = fund_xx.iloc[fund_index1,2]
        fund_idxname.append(fund_day)
        m = 0
        while m < howlong :
                fund_index1 =  fund_index1 + daynum
                fund_day = fund_xx.iloc[fund_index1,2]
                fund_idxname.append(fund_day)
                m = m + 1

        fund_idxname.pop()
        fund_idxname.insert(0,'最新日期')
        fund_idxname.insert(0,'基金简称')
        fund_idxname.insert(0,'基金代码')
        fundrate_result = pd.DataFrame(columns=fund_idxname)
        print(fundrate_result)
        #保存基金涨幅的索引号
        j = 0 
        #开始计算每个基金的涨幅
        for fund in list1:
            fund_xx=fundlist[fundlist['基金代码']==fund]
            fund_indexs = []
            fund_rates = []
            if len(fund_xx.index) < min_count:
                continue
            m=0
            #单独过滤出来的某只基金的pd，索引编号都是0开始，索引名称并不一定是顺序的（因为多线程抓取的原因）,所以下面都用iloc
            fund_index1 = 0
            fund_indexs.append(fund_index1)
            while m < howlong :
                fund_index1 =  fund_index1 + daynum
                fund_indexs.append(fund_index1)
                #列从0开始，3表示第4列，是基金净值
                first_value = fund_xx.iloc[fund_indexs[m],3]
                second_value = fund_xx.iloc[fund_indexs[m+1],3]
                first_value = float(first_value)
                second_value = float(second_value)
                first_rate =  round((first_value - second_value)/second_value * 100,2)
                fund_rates.append(first_rate)
                m = m + 1
            name = fund_xx.iloc[fund_indexs[m],1]
            fund_rates.insert(0,lasttrans_day)
            fund_rates.insert(0,name)
            fund_rates.insert(0,fund)
            print(fund_rates)
            fundrate_result.loc[j]=fund_rates
            j = j + 1

    #按照datetype=CD来处理
    else:
        print("I don't finish this part")
        #if type=="week":
        
        #按照type=="month"来处理
        #else:
     
    return fundrate_result


# In[126]:


#对排序获得的集合数据pddata, 进行画图并保存图片名为picname， wz 是画图数据的列的位置， *col是从pddata中选择哪些列的列表
def pic_execl(pddata,picname,st,wz,*col):
    #pdata = pd.DataFrame(pddata,columns=['序号', '基金代码', '基金简称', '日期', '自选', '最近1月', '最近2-3月', '最近4-6月', '最近7-12月', '最近1-2年', '最近2-3年', '从前'])
    if len(col) > 0 :
        pdata = pd.DataFrame(pddata,columns=col)
    else:
        pdata = pddata
    # print ("--pdata---")
    # print(pdata)
    #横坐标是代表周期的列名，从0开始，这里是从4开始，到传入的位置截至。 如上面的列子是从："自选"开始
    xtime = pdata.columns.values[st:wz]
    # print("---xtime---")
    # print(xtime)
    #取每一行的数据组成曲线图，并将"基金代码，基金简称" 作为曲线名称
    for index, row in pdata.iterrows():
        # print(list(row)[5:-1])
        plt.plot(xtime, list(row)[st:wz], '.-', label=list(row)[0:2])
    plt.xticks(xtime)
    plt.xticks(rotation=90)
    fig, ax = plt.subplots()
    plt.figure(1)
    plt.xlabel('周期')
    plt.ylabel('涨幅百分比')
    plt.legend(loc='center left',bbox_to_anchor=(1.0,0.5))
    #fig.subplots_adjust(right=0.6)
    plt.savefig(time.strftime("%Y%m%d", time.localtime()) + "_" + picname + ".png",dpi=600,bbox_inches='tight')
    plt.show()


# In[127]:


'''
通常用于我自己持仓基金的排序取交集，因为我持仓的基金抓取数据会抓最近半年多的，所以周按最近26个周，月按最近6个月，分别进行每个周期的排序
将传入的计算好的包含基金rate的 pd， 进行按rate排序，取top_n。 然后再将每列取的top_n的 pd， 取交集
'''
def myfund_rate_sort(fund_ratelist,type,top_n):
    #定义一个动态变量
    names = locals()
    
    if type=="week":
        sortlist=26       
    if type=="month":
        sortlist=6

    m=0
    #fund_ratelist的前3列为：基金代码	基金简称	最新日期， 所以从第4列(即下标3）开始排序，并取前top_n 行的 pd
    n=3

    for i in range(sortlist):  
        #将每列的top_n, 然后付给动态变量
        names['s'+ str(i)] = fund_ratelist.sort_values(by=fund_ratelist.columns.values[n],ascending=False).head(top_n)
        print("*** fund_top ***")
        #展示动态变量中的pd
        print(names.get('s'+ str(i)), end = '\n' )
        m = m + 1
        n = n + 1
   
    #将动态变量中包含每个rate列top_n的pd，求交集   
    j=0
    k=1
    inner_sort = pd.merge(names.get('s'+str(j)), names.get('s'+str(k)), how='inner')
    k = k + 1
    while k < sortlist:
        inner_sort = pd.merge(inner_sort, names.get('s'+ str(k)), how='inner')
        k = k + 1
    print("****** sort innner pd *****")
    print(inner_sort)
    return inner_sort


# In[140]:


'''
将传入的计算好的包含基金rate的 pd， 按照最近period_n个周期，对每个周期进行按rate排序，取top_n。 然后再将这些周期的top_n的 pd取交集
'''
def fund_rate_sort(fund_ratelist,period_n,top_n):
    #定义一个动态变量
    names = locals()
    
    #要排序几个周期
    sortlist=period_n

    m=0
    #fund_ratelist的前3列为：基金代码	基金简称	最新日期， 所以从第4列(即下标3）开始排序，并取前top_n 行的 pd
    n=3

    for i in range(sortlist):  
        #将每列的top_n, 然后付给动态变量
        names['s'+ str(i)] = fund_ratelist.sort_values(by=fund_ratelist.columns.values[n],ascending=False).head(top_n)
        print(str(i) + "*** fund_top ***")
        #展示动态变量中的pd
        print(names.get('s'+ str(i)), end = '\n' )
        m = m + 1
        n = n + 1
   
    #将动态变量中包含每个rate列top_n的pd，求交集   
    j=0
    k=1
    inner_sort = pd.merge(names.get('s'+str(j)), names.get('s'+str(k)), how='inner')
    k = k + 1
    while k < sortlist:
        inner_sort = pd.merge(inner_sort, names.get('s'+ str(k)), how='inner')
        k = k + 1
    print("****** sort innner pd *****")
    print(inner_sort)
    return inner_sort


# In[142]:



def main():
    url="http://fundf10.eastmoney.com/jjjz_000066.html"
    

if __name__ == '__main__':
    
    #基金详细交易信息获取后，保存到该名称的execl中
    sourceexecl="fund_mingxi"
    #对获取的基金信息进行编辑，增加10列后保存的execl名称
    editexecl="fundedit"
    #对月份进行排序取交集保存的execl名称
    sortexecl="fund_month_sort"
    #对年份进行排序取交集保存的execl名称
    sort1execl="find_year_sort"
    #对基金明细数据进行保存的execl名称
    allfunds_detailexecl="allfund_trans_detail"

    #对我持有基金明细数据进行保存的execl名称
    myfunds_detailexecl="myfunds_trans_detail"

    #获取当天日期
    cur_day=time.strftime("%Y-%m-%d", time.localtime())

    my_funds = ["5275", "162605", "110011", "270050", "83", "519674", "486001", "727","210008"]

    #调用函数，打开url，抓取最后一个交易日日期
    url="http://fundf10.eastmoney.com/jjjz_000066.html"
    driver = open_url(url)
    lasttrans_day=driver.find_element_by_xpath('//*[@id="jztable"]/table/tbody/tr[1]/td[1]').text

    #读取最全的基金的信息，并转换成pd
    edit_funds=pd.read_excel(r"C:\工作\pychram\111\fund_get\2020-07-11\fundedit.xls")
    #创建一个pd，只包含两列数据
    all_funds=edit_funds[['基金代码','基金简称']]

    #从全量的基金pd中，获取我持有列表的基金代码和基金名称的pd
    myfunds_numname = fund_num_name(all_funds,*my_funds)

    #对我持有的基金代码和基金名称pd，抓取每只基金的交易明细， 保存为myfunds_detailexecl的execl文件， 并返回一个我持有基金明细数据的pd
    my_f=getfund_mingxi(myfunds_numname,myfunds_detailexecl)
    
    #退出chrome
    driver.quit()

    #读取detailexecl的execl文件， 并去掉首列的内容，获取个全量明细的pd
    #xx=pd.read_excel(myfunds_detailexecl+".xlsx", usecols=[1,2,3,4,5,6])

    #读取我持有基金的基金明细的pd，按照“week/month" 来计算涨幅，日期按照间隔5个工作日来算， 返回我持有基金最近26周的每周涨幅
    my_w_rate1 = fund_rate(my_f,"week","WD",5)

    #读取我持有基金的基金明细的pd，按照“week/month" 来计算涨幅，日期按照间隔23个工作日来算，返回我持有基金最近6月的每月涨幅
    my_m_rate1 = fund_rate(my_f,"month","WD",23)

    #对按月的涨幅进行画图
    pic_execl(my_m_rate1,"6month",3,9)
    #对按周的涨幅进行画图
    pic_execl(my_w_rate1,"26week",3,29)
    
    #对按月涨幅的全量数据，取每个月的涨幅前五，取交集
    my_sort_inner_rate = myfund_rate_sort(my_m_rate1,"month",5)
    try:
        pic_execl(my_sort_inner_rate,"my_top5_6month",3,9)
    except ValueError:
        print("Can't find inner funds")
    else:
        pass


# In[70]:


    filenum = 1
    while len(all_funds) !=0 :
        #计划用3个线程进行明细数据抓取，这里定义每个线程抓取哪些基金
        count = int(len(all_funds)/3)
        funds_part1=all_funds.iloc[:count]
        funds_part2=all_funds.iloc[count:count*2]
        funds_part3=all_funds.iloc[count*2:]

        #启动三个浏览器，进行多线程程的抓取
        driver3 = webdriver.Chrome(executable_path=chrome_driver)
        driver1 = webdriver.Chrome(executable_path=chrome_driver1)
        driver2 = webdriver.Chrome(executable_path=chrome_driver2)

        '''
        启动第一个线程，该线程使用驱动器1，将参数传给多线程抓取函数，处理第一部分基金，抓取3页（60个工作日的基金明细），
        多线程处理抓取的数据结果会放置到全局变量的tmp_funds1 的pd中
        '''
        print("开始多线程抓取时间：")
        print (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
        t1 = threading.Thread(target=getfund_mingxi_mt , args=(funds_part1,driver1,3))
        t1.start()
        t2 = threading.Thread(target=getfund_mingxi_mt , args=(funds_part2,driver2,3))
        t2.start()
        t3 = threading.Thread(target=getfund_mingxi_mt , args=(funds_part3,driver3,3))
        t3.start()
        t1.join()
        t2.join()
        t3.join()
        print("结束多线程抓取时间：")
        print (time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

        #上面可能会发生异常退出，并没有抓取完成所有的基金详情，先退出chrome
        driver3.quit()
        driver1.quit()
        driver2.quit()

        #将已经抓取的基金去重，然后基金代码由字符串转为int，因为all_funds中基金代码都是int
        cur_getfunds=tmp_funds1[['基金代码','基金简称']].drop_duplicates()
        cur_getfunds['基金代码'] = pd.to_numeric(cur_getfunds['基金代码'])

        #将all_funds和cur_getfunds取补集，即all_funds中不包含cur_getfunds的部分
        all_funds = all_funds.append(cur_getfunds)
        all_funds = all_funds.append(cur_getfunds)
        all_funds = all_funds.drop_duplicates(subset=['基金代码','基金简称'],keep=False)

        #将重复抓取的内容去掉
        tmp_funds1=tmp_funds1.drop_duplicates()

        filenum = filenum + 1 
        #将多线程抓取的基金明细信息保存到execl中，并且不保存index
        tmp_funds1.to_excel("all_fund_" + str(cur_day) + "_" + str(filenum) + ".xlsx",index=False)


    # In[132]:


    #将所有基金的明细，按周计算增幅，每周5个工作日， 取最近6周每周的涨幅
    all_w_rate1 = fund_rate_mt(tmp_funds1,"week","WD",5,6)
    #将所有基金周涨幅，取最近3周，每周top500，求交集
    all_week_sort_rate = fund_rate_sort(all_w_rate1,3,500)


    # In[ ]:


    #将所有基金的明细，按月计算增幅，每周22个工作日， 取最近2月每周的涨幅
    all_m_rate1 = fund_rate_mt(tmp_funds1,"week","WD",22,2)
    #将所有基金月涨幅，取最近2月，每月top200，求交集
    all_month_sort_rate = fund_rate_sort(all_m_rate1,2,200)


    # In[ ]:


    #对所有基金明细交集的内容，画图
    pic_execl(all_week_sort_rate,"all_top500_3week",3,6)
    pic_execl(all_month_sort_rate,"all_top200_2month",3,5)




