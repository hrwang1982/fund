# encoding: UTF-8
import os
import sys
import re
import xlwt
import xlrd
from xlutils.copy import copy


def main():
    fundxls=xlrd.open_workbook("fund.xls")
    sheets=fundxls.sheet_names()
    worksheet = fundxls.sheet_by_name(sheets[0])
    rows_old=worksheet.nrows
    cols_old=worksheet.ncols
    print("---rows_old----")
    print (rows_old)
    print ("---cols_old----")
    print (cols_old)
    new_fundxls=copy(fundxls)
    new_worksheet=new_fundxls.get_sheet(0)

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


    for i in range(1,rows_old):
        #custome duration
        if worksheet.cell_value(i, 17).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old, float(worksheet.cell_value(i, 17).split("%")[0]))
        #last month
        if worksheet.cell_value(i,9).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 1, float(worksheet.cell_value(i, 9).split("%")[0]))
        #last 2month-3month
        if worksheet.cell_value(i, 10).split("%")[0] not in "----" and worksheet.cell_value(i, 9).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 10).split("%")[0]))
            m2_last=float(worksheet.cell_value(i,10).split("%")[0])-float(worksheet.cell_value(i,9).split("%")[0])
            new_worksheet.write(i,cols_old + 2,float(m2_last))
        #last 3 month
        if worksheet.cell_value(i,10).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 3, float(worksheet.cell_value(i, 10).split("%")[0]))
        #last 4month-6month
        if worksheet.cell_value(i, 11).split("%")[0] not in "----" and worksheet.cell_value(i, 10).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 11).split("%")[0]))
            m3_last=float(worksheet.cell_value(i,11).split("%")[0])-float(worksheet.cell_value(i,10).split("%")[0])
            new_worksheet.write(i,cols_old + 4,float(m3_last))
        #last 6 month
        if worksheet.cell_value(i,11).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 5, float(worksheet.cell_value(i, 11).split("%")[0]))
        # last 7month-12month
        if worksheet.cell_value(i, 12).split("%")[0] not in "----" and worksheet.cell_value(i, 11).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 12).split("%")[0]))
            m4_last=float(worksheet.cell_value(i,12).split("%")[0])-float(worksheet.cell_value(i,11).split("%")[0])
            new_worksheet.write(i,cols_old + 6,float(m4_last))
        #duration now-1year
        if worksheet.cell_value(i, 12).split("%")[0] not in "----":
            new_worksheet.write(i, cols_old + 7, float(worksheet.cell_value(i, 12).split("%")[0]))
        #duration 1year-2year
        if worksheet.cell_value(i, 13).split("%")[0] not in "----" and worksheet.cell_value(i, 12).split("%")[0] not in "----":
            print(float(worksheet.cell_value(i, 13).split("%")[0]))
            y2_last=float(worksheet.cell_value(i,13).split("%")[0])-float(worksheet.cell_value(i,12).split("%")[0])
            new_worksheet.write(i, cols_old + 8, float(y2_last))
        #duration 2year-3year
        if worksheet.cell_value(i, 14).split("%")[0] not in "----" and worksheet.cell_value(i, 13).split("%")[0] not in "----":
            y3_last = float(worksheet.cell_value(i, 14).split("%")[0])- float(worksheet.cell_value(i, 13).split("%")[0])
            new_worksheet.write(i, cols_old + 9, float(y3_last))
        #duration 3year-old
        if worksheet.cell_value(i, 16).split("%")[0] not in "----" and worksheet.cell_value(i, 14).split("%")[0] not in "----":
            y3_old = float(worksheet.cell_value(i, 16).split("%")[0]) - float(worksheet.cell_value(i, 14).split("%")[0])
            new_worksheet.write(i, cols_old + 10, float(y3_old))


    # for i in range(0, new_rows):
    #     for j in range(0,new_cols):
    #         print(new_worksheet.cell_value(i,j), "\t", end="")
    # print()
    new_fundxls.save("fund1.xls")


    fund1xls=xlrd.open_workbook("fund1.xls")
    sheets1=fund1xls.sheet_names()
    worksheet1 = fund1xls.sheet_by_name(sheets1[0])
    new_rows=worksheet1.nrows
    new_cols=worksheet1.ncols
    print("---rows_new----")
    print (new_rows)
    print ("---cols_new----")
    print (new_cols)

if __name__ == '__main__':
     main()
