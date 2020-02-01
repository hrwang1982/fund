# encoding: UTF-8
import os
import urllib.request
import json
import urllib.parse
import sys
import traceback
import requests
import re
from selenium import webdriver
import xlwt

chrome_driver=r'C:\工作\pychram\111\chromedriver.exe'


def write_to_file(content):
    with open('xiaoxi.txt', 'a', encoding='utf-8')as f:
        # print(type(json.dumps(content)))
        f.write(json.dumps(content,ensure_ascii=False))

def main():
    #混合型
    #url = "http://fund.eastmoney.com/data/fundranking.html#thh;c0;r;s1nzf;pn10000;ddesc;qsd20180501;qed20181231"
    #全部
    url = "http://fund.eastmoney.com/data/fundranking.html#tall;c0;r;s1nzf;pn10000;ddesc;qsd20180501;qed20181231;qdii;zq;gg;gzbd;gzfs;bbzt;sfbb"
    driver=webdriver.Chrome(executable_path=chrome_driver)

    driver.get(url)
    driver.implicitly_wait(20)
    html_text=driver.find_element_by_xpath('//div[@class="mainFrame"][7]/div[@class="dbtable"]/table[@id="dbtable"]/tbody').find_elements_by_tag_name("tr")
    list=[]

    execlfile=xlwt.Workbook()
    sheet=execlfile.add_sheet('fund',cell_overwrite_ok=True)
    n = 0
    for thead in ("比较","序号","基金代码","基金简称","日期","单位净值","累计净值","日增长率","近1周","近1月","近3月","近6月","近1年","近2年","近3年","今年来","成立来","自定义","手续费"):
        sheet.write(0,n,thead)
        n+=1

    m=1
    for i in html_text:
        j=i.find_elements_by_tag_name("td")
        n=0
        for item in j:
            text=item.text
            sheet.write(m,n,text)
            print (text)
            n+=1
        m+=1
    execlfile.save("fund.xls")
    driver.quit()


if __name__ == '__main__':

    main()