# -*- encoding:UTF-8 -*-
import sitecustomize
import xlrd
import xlwt
import os
from xlutils.copy import copy
import time

def readXls():
    print unicode('请将需要统计的报表放在该目录的data文件夹下,并从0.xls开始顺序命名', 'utf-8')
    print unicode('请输入待汇总的表格份数','utf-8')
    num=input()
    sheet = {}
    for each in range(num):
        data = xlrd.open_workbook('data/'+str(each)+'.xls')
        table = data.sheets()[0]
        list=[]
        cell_B1 = table.cell(0,4).value
        list.append(cell_B1)
        morning = table.cell(34,5).value
        noon = table.cell(34,6).value
        night = table.cell(34,7).value
        twelve = table.cell(35,8).value
        bus = table.cell(34,9).value
        ill = table.cell(34,10).value
        kuang = table.cell(34,11).value
        list.append(morning)
        list.append(noon)
        list.append(night)
        list.append(twelve)
        list.append(bus)
        list.append(ill)
        list.append(kuang)
        base = table.cell(12,18).value
        super = table.cell(12,19).value
        list.append(base)
        list.append(super)
        sheet[list[0]]=list
    return sheet

def writeXls(sheet):
    file = xlwt.Workbook()
    table = file.add_sheet('1')
    i=0
    for list in sheet.values():
        j=0
        for each in list:
            table.write(i,j,each)
            j+=1
        i+=1
    file.save('result.xls')

if __name__=='__main__':
    sheet = {}
    sheet = readXls()
    writeXls(sheet)
    print "run successfully!"
    os.system("pause")