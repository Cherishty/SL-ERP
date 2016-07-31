# -*- encoding:UTF-8 -*-
import sitecustomize
import xlrd
import xlwt
import os
from xlutils.copy import copy
import time


def readXls():
    print('------------------------------------')
    print unicode('自络筒工序 工资计算v1.3', 'utf-8')  # decode via unicode function
    print unicode('请将需要统计的报表放在该目录的data文件夹下,并从0.xls开始顺序命名', 'utf-8')
    print unicode('请输入待汇总的表格份数', 'utf-8')
    companyList = {}

    num = raw_input()
    while num.isalpha():
        print unicode('输入错误，请重新输入合法数字', 'utf-8')
        num = raw_input()

    for each in range(int(num)):
        try:
            data = xlrd.open_workbook('data/' + str(each) + '.xls')
        except:
            print unicode('没有找到该目录或该文件，请检查您的路径或命名\n\n', 'utf-8')
            print '------------------------------------\n\n'
            return
        else:
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
    file = xlwt.Workbook(encoding='utf-8')  # encoding='utf-8' to support chinese output
    table = file.add_sheet('1')
    i = 0
    k = 0
    title = ['公司名称','金额']
    for each in title:
        table.write(i, k, each)
        k += 1
    i += 1

    for key in sheet:
        table.write(i, 0, key)
        table.write(i, 1, sheet[key])
        i += 1
    file.save('result.xls')


if __name__ == '__main__':
    sheet = {}
    while not sheet:
        sheet = readXls()
    writeXls(sheet)
    print "run successfully!"
    os.system("pause")
