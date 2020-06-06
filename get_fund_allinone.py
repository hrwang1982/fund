# encoding: UTF-8
from selenium import webdriver
import xlwt
import xlrd
from xlutils.copy import copy
import pandas as pd
import matplotlib.pyplot as plt
import time

#使用的chrome驱动器
chrome_driver=r'C:\工作\pychram\111\chromedriver.exe'

# 这两行代码解决 plt 中文显示的问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

#打开chrome浏览器，打开指定的url，并返回driver
def open_url(url):
    driver = webdriver.Chrome(executable_path=chrome_driver)
    driver.get(url)
    driver.implicitly_wait(20)
    return driver

#从指定的url中，反馈基金的内容
def get_table(webdriver):
    html_text = webdriver.find_element_by_xpath('//div[@class="mainFrame"][7]/div[@class="dbtable"]/table[@id="dbtable"]/tbody').find_elements_by_tag_name("tr")
    return html_text

#将从网页中获取的基金内容"context"，写入 "filename.xls" 的表格中，表中的列如下面的内容：
def write_excle(filename,context):
    execlfile = xlwt.Workbook()
    sheet = execlfile.add_sheet(filename, cell_overwrite_ok=True)
    #写execl标题列到execl文件中，默认行和列都是从0开始
    n = 0
    for thead in ("比较", "序号", "基金代码", "基金简称", "日期", "单位净值", "累计净值", "日增长率", "近1周", "近1月", "近3月", "近6月", "近1年", "近2年", "近3年", "今年来","成立来", "自定义", "手续费"):
        sheet.write(0, n, thead)
        n += 1

    #将context表格中的内容按行写入execl文件
    m = 1
    for line in context:
        clos = line.find_elements_by_tag_name("td")
        n = 0
        for item in clos:
            text = item.text
            sheet.write(m, n, text)
            print(text)
            n += 1
        m += 1
    execlfile.save(filename+".xls")

'''
读取保存了基金信息的"readexecl.xls" 表格， 并创建一个新的"editexecl.xls"表格，增加10列（这10列都是数值型，且不带百分号）：
自选，最近1月，最近2-3月，最近3月，最近4-6月，最近6月，最近7-12月，最近1年，最近1-2年，最近2-3年，从前
'''
def edit_execl(readexecl,editexecl):
    #读取保存基金信息的execl表格
    fundxls = xlrd.open_workbook(readexecl+".xls")
    sheets = fundxls.sheet_names()
    worksheet = fundxls.sheet_by_name(sheets[0])

    #打印该表格的行数和列数
    rows_old = worksheet.nrows
    cols_old = worksheet.ncols
    print("---rows_old----")
    print(rows_old)
    print("---cols_old----")
    print(cols_old)

    #拷贝一个新的execl表格
    new_fundxls = copy(fundxls)
    new_worksheet = new_fundxls.get_sheet(0)

    #新execl表格，在标题行，增加10列的列标题
    new_worksheet.write(0, cols_old,"自选")
    new_worksheet.write(0, cols_old + 1, "最近1月")
    new_worksheet.write(0, cols_old + 2, "最近2-3月")
    new_worksheet.write(0, cols_old + 3, "最近3月")
    new_worksheet.write(0, cols_old + 4, "最近4-6月")
    new_worksheet.write(0, cols_old + 5, "最近6月")
    new_worksheet.write(0, cols_old + 6, "最近7-12月")
    new_worksheet.write(0, cols_old + 7, "最近1年")
    new_worksheet.write(0, cols_old + 8, "最近1-2年")
    new_worksheet.write(0, cols_old + 9, "最近2-3年")
    new_worksheet.write(0, cols_old + 10, "从前")

    #新execl表格行从0开始（0代表标题行），循环对每一行计算新增的10列的内容
    for i in range(1,rows_old):
        #custome duration  自选，(列是0开始），17列是"自定义" 列
        if worksheet.cell_value(i, 17).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old, float(worksheet.cell_value(i, 17).split("%")[0]))
        #last month  最近1月 ，（列是0开始），9列是"近1月" 列
        if worksheet.cell_value(i,9).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 1, float(worksheet.cell_value(i, 9).split("%")[0]))
        #last 2month-3month 最近2-3月，（列是0开始），10列是"近3月" 列
        if worksheet.cell_value(i, 10).split("%")[0] not in "----" and worksheet.cell_value(i, 9).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 10).split("%")[0]))
            m2_last=float(worksheet.cell_value(i,10).split("%")[0])-float(worksheet.cell_value(i,9).split("%")[0])
            new_worksheet.write(i,cols_old + 2,float(m2_last))
        #last 3 month 最近3月，（列是0开始），10列是"近3月" 列
        if worksheet.cell_value(i,10).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 3, float(worksheet.cell_value(i, 10).split("%")[0]))
        #last 4month-6month 最近4-6月，（列是0开始），11列是"近6月" 列
        if worksheet.cell_value(i, 11).split("%")[0] not in "----" and worksheet.cell_value(i, 10).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 11).split("%")[0]))
            m3_last=float(worksheet.cell_value(i,11).split("%")[0])-float(worksheet.cell_value(i,10).split("%")[0])
            new_worksheet.write(i,cols_old + 4,float(m3_last))
        #last 6 month 最近6月，（列是0开始），11列是"近6月" 列
        if worksheet.cell_value(i,11).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 5, float(worksheet.cell_value(i, 11).split("%")[0]))
        # last 7month-12month 最近7-12月，（列是0开始），12列是"近1年" 列
        if worksheet.cell_value(i, 12).split("%")[0] not in "----" and worksheet.cell_value(i, 11).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 12).split("%")[0]))
            m4_last=float(worksheet.cell_value(i,12).split("%")[0])-float(worksheet.cell_value(i,11).split("%")[0])
            new_worksheet.write(i,cols_old + 6,float(m4_last))
        #duration now-1year 最近1年，（列是0开始），12列是"近1年" 列
        if worksheet.cell_value(i, 12).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 7, float(worksheet.cell_value(i, 12).split("%")[0]))
        #duration 1year-2year 最近1-2年，（列是0开始），13列是"近2年" 列
        if worksheet.cell_value(i, 13).split("%")[0] not in "----" and worksheet.cell_value(i, 12).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 13).split("%")[0]))
            y2_last=float(worksheet.cell_value(i,13).split("%")[0])-float(worksheet.cell_value(i,12).split("%")[0])
            new_worksheet.write(i, cols_old + 8, float(y2_last))
        #duration 2year-3year 最近2-3年，（列是0开始），14列是"近3年" 列
        if worksheet.cell_value(i, 14).split("%")[0] not in "----" and worksheet.cell_value(i, 13).split("%")[0] not in "----":
            y3_last = float(worksheet.cell_value(i, 14).split("%")[0])- float(worksheet.cell_value(i, 13).split("%")[0])
            new_worksheet.write(i, cols_old + 9, float(y3_last))
        #duration 3year-old 从前，（列是0开始），16列是"成立来" 列
        if worksheet.cell_value(i, 16).split("%")[0] not in "----" and worksheet.cell_value(i, 14).split("%")[0] not in "----":
            y3_old = float(worksheet.cell_value(i, 16).split("%")[0]) - float(worksheet.cell_value(i, 14).split("%")[0])
            new_worksheet.write(i, cols_old + 10, float(y3_old))

    #保存编辑好的execl文件
    new_fundxls.save(editexecl+".xls")

    fund1xls = xlrd.open_workbook(editexecl+".xls")
    sheets1 = fund1xls.sheet_names()
    worksheet1 = fund1xls.sheet_by_name(sheets1[0])

    # 打印编辑好的表格的行数和列数
    new_rows = worksheet1.nrows
    new_cols = worksheet1.ncols
    print("---rows_new----")
    print(new_rows)
    print("---cols_new----")
    print(new_cols)

#在编辑好的"sourceexecl"表格中查找我持仓的基金的数据,返回该数据集，并用于画图
def get_myfunds(sourceexecl,myfunds_list,*col):
    #读取处理过的数据文件
    mydata = pd.read_excel(sourceexecl+'.xls', sheet_name=0)
    mydata1 = pd.DataFrame(mydata)
    pdata=pd.DataFrame(mydata1, columns=col)
    pdata1=pd.DataFrame(columns=col)
    for my_fund in myfunds_list:
        for row in pdata.iterrows():
            if my_fund == str(row[1]):
                row1 = dict(row)
                pdata1 = pdata1.append(row1, ignore_index=True)
    return pdata1

#对编辑好的"sourceexecl"表格，选择5列, 进行TOPN排序，并取交集,并保存为"sortexecl"表格。
def sort_execl(sourceexecl, sortexecl, first, second, third, forth, fifth, topn=200):
    #读取增加了10列的execl表格
    data = pd.read_excel(sourceexecl+".xls", sheet_name=0)
    # data1=pd.DataFrame(data,columns=['序号','基金代码','基金简称','日期','自选','1年','2年','3年','从前'])
    data1 = pd.DataFrame(data)

    #对传入的第1列"first"(first为列名)进行倒排，取收益率最大的topn
    df0 = data1.sort_values(first, ascending=False).head(topn)
    # 对传入的第2列"second"(second为列名)进行倒排，取收益率最大的topn
    df1 = data1.sort_values(second, ascending=False).head(topn)
    df2 = data1.sort_values(third, ascending=False).head(topn)
    df3 = data1.sort_values(forth, ascending=False).head(topn)
    df4 = data1.sort_values(fifth, ascending=False).head(topn)

    #对传入的第1列的收益率最大的topn和传入的第2列的收益率最大的topn   取交集
    df01 = pd.merge(df0, df1, how='inner')
    #第1列和2列的交集， 再和第3列的收益率最大的topn   取交集
    df012 = pd.merge(df01, df2, how='inner')
    ##第1列、2列、3列的交集， 再和第4列的收益率最大的topn   取交集
    df0123 = pd.merge(df012, df3, how='inner')
    df01234 = pd.merge(df0123, df4, how='inner' )

    #如果5个排序的列，存在交集
    if df01234.shape[0] >= 1 :
        df01234.to_excel(sortexecl+".xls", sheet_name=sortexecl, encoding='utf-8')
        print(first + ", " + second + ", " + third + ", " + forth + ", " + fifth + ", " + "存在交集")
        return df01234
    #如果5个排序的列不存在交集，但是前4列存在交集
    elif df01234.shape[0] < 1 and df0123.shape[0] >= 1 :
        df0123.to_excel(sortexecl+".xls", sheet_name=sortexecl, encoding='utf-8')
        print(first + ", " + second + ", " + third + ", " + forth + ", " + "存在交集")
        return df0123
    elif df01234.shape[0] < 1 and df0123.shape[0] < 1 and df012.shape[0] >=1 :
        df012.to_excel(sortexecl+".xls", sheet_name=sortexecl, encoding='utf-8')
        print(first + ", " + second + ", " + third + ", " + "存在交集")
        return df012
    elif df01234.shape[0] < 1 and df0123.shape[0] < 1 and df012.shape[0] <1 and df01.shape[0] >=1 :
        df01.to_excel(sortexecl+".xls", sheet_name=sortexecl, encoding='utf-8')
        print(first + ", " + second + ", " + "存在交集")
        return df01
    else :
        print(u"没有交集")

#对排序获得的集合数据pddata, 进行画图并保存图片名为picname， wz 是画图数据的列的位置， *col是从pddata中选择哪些列的列表
def pic_execl(pddata,picname,wz,*col):
    #pdata = pd.DataFrame(pddata,columns=['序号', '基金代码', '基金简称', '日期', '自选', '最近1月', '最近2-3月', '最近4-6月', '最近7-12月', '最近1-2年', '最近2-3年', '从前'])
    pdata = pd.DataFrame(pddata,columns=col)

    #横坐标是代表周期的列名，从0开始，这里是从4开始，到传入的位置截至。 如上面的列子是从："自选"开始
    xtime = pdata.columns.values[4:wz]

    #取每一行的数据组成曲线图，并将"基金代码，基金简称" 作为曲线名称
    for index, row in pdata.iterrows():
        # print(list(row)[5:-1])
        plt.plot(xtime, list(row)[4:wz], '.-', label=list(row)[1:3])
    plt.xticks(xtime)
    plt.xlabel('周期')
    plt.ylabel('涨幅百分比')
    plt.legend()
    plt.savefig(time.strftime("%Y%m%d", time.localtime()) + "_" + picname + ".png")
    plt.show()


def main():
    #混合型
    #url = "http://fund.eastmoney.com/data/fundranking.html#thh;c0;r;s1nzf;pn10000;ddesc;qsd20180501;qed20181231"
    #全部
    url = "http://fund.eastmoney.com/data/fundranking.html#tall;c0;r;s1nzf;pn10000;ddesc;qsd20180501;qed20181231;qdii;zq;gg;gzbd;gzfs;bbzt;sfbb"


if __name__ == '__main__':
    #获取基金信息的url，其中的20180501 - 20181231 是自定义的周期时间
    url = "http://fund.eastmoney.com/data/fundranking.html#tall;c0;r;s1nzf;pn10000;ddesc;qsd20180501;qed20181231;qdii;zq;gg;gzbd;gzfs;bbzt;sfbb"
    #基金信息获取后，保存到该名称的execl中
    sourceexecl="fundsave"
    #对获取的基金信息进行编辑，增加10列后保存的execl名称
    editexecl="fundedit"
    #对月份进行排序取交集保存的execl名称
    sortexecl="fund_month_sort"
    #对年份进行排序取交集保存的execl名称
    sort1execl="find_year_sort"

    #我目前持仓的基金列表
    my_funds = ("005275", "162605", "001076", "110011", "270050", "000083", "519674", "486001")

    #调用函数，打开url
    driver=open_url(url)

    #调用函数将基金信息保存到table_context中
    table_context=get_table(driver)

    #调用函数，将基金信息保存到execl中
    write_excle(sourceexecl,table_context)

    #退出chrome
    driver.quit()

    #编辑基金信息，增加10列
    edit_execl(sourceexecl,editexecl)

    #按照月进行交集获取，并画图
    msort=sort_execl(editexecl,sortexecl,'最近1月','最近2-3月','最近4-6月','最近7-12月','最近1年',200)
    mlist=['序号', '基金代码', '基金简称', '日期', '最近1月','最近2-3月','最近4-6月','最近7-12月','最近1年']
    pic_execl(msort,"monthsort",-1,*mlist)

    #按照年进行交集获取，并画图
    ysort=sort_execl(editexecl, sort1execl, '自选', '最近1年', '最近1-2年', '最近2-3年', '从前', 1000)
    ylist=['序号', '基金代码', '基金简称', '日期', '自选','最近1年','最近1-2年','最近2-3年','从前']
    pic_execl(ysort, "yearsort",9 , *ylist)

    #将我持仓的基金的最新数据获取，画图
    my_cur_funds=get_myfunds(editexecl, my_funds, *mlist)
    pic_execl(my_cur_funds, "my_funds", -1, *mlist)
