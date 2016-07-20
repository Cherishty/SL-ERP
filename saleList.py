# -*- encoding:UTF-8 -*-
import sitecustomize
import xlrd
import xlwt
import os
from xlutils.copy import copy
import time


def readXls():
    print unicode('请将需要统计的报表放在该目录的data文件夹下,并从0.xls开始顺序命名', 'utf-8')
    print unicode('请输入待汇总的表格份数', 'utf-8')
    num = input()
    companyList = {}
    for each in range(num):
        data = xlrd.open_workbook('data/' + str(each) + '.xls')
        table = data.sheets()[0]
        row = table.nrows
        for eachRow in range(2, row):
            companyName = table.cell(eachRow, 2).value
            if not companyName:
                break
            money = table.cell(eachRow, 6).value
            if companyList.has_key(companyName):
                companyList[companyName] = companyList[companyName] + float(money)
            else:
                companyList[companyName] = float(money)
    return companyList


def writeXls(sheet):
    file = xlwt.Workbook()
    table = file.add_sheet('1')
    i = 0
    for key in sheet:
        table.write(i, 0, key)
        table.write(i, 1, sheet[key])
        i += 1
    file.save('result.xls')

if __name__ == '__main__':
    sheet = {}
    sheet = readXls()
    writeXls(sheet)
    print "run successfully!"
    os.system("pause")
