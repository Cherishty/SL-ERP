import xlrd
import glob
import xlwt
from xlutils.copy import copy
import XLFormat
import xml.etree.ElementTree as ET
import os
import sys

def ReadXls(result):
    print('\n------------------------------------')
    print('公司月度工资汇总 v1.2 \n')
    print ("启动中\n\n")
    print('请参阅使用说明，并将待处理的文件放在data目录下\n')
    print("导出的模板在sheet中，如需要修改格式，请联系我\n")

    personInBnak = {}
    personInReal = {}
    departmentInfo = []
    file =glob.glob('data/*.xls')
    print(file)
    try:
        data = xlrd.open_workbook(file[0])
    except:
        print('没有找到待读取的%s，操作失败！请检查您的文件\n\n' %file[0])
        return

    print('%s读取成功，正在处理......' %file[0])

    # record staffs' salary detail in every department
    for sheet in range(0,7):
        table = data.sheets()[sheet]
        currentStaff=0
        currentInsurance=0 #people who join Social Insurance
        rows = table.nrows
        departmentName=table.cell(0, 0).value.split()[0]
        for row in range(4, rows):
            staffNanme = table.cell(row, 1).value.replace(' ', '')

            if(staffNanme=="合计"):
                fkSumMoney=table.cell(row, 20).value
                insureMoney=table.cell(row, 21).value
                tax=table.cell(row, 22).value
                sumBankMoney=table.cell(row, 23).value
                sumRealMoney=table.cell(row, 24).value
                totalMoney=table.cell(row, 25).value
                departmentInfo.append([departmentName,currentStaff,fkSumMoney,currentInsurance,insureMoney,0,tax,sumBankMoney,sumRealMoney,totalMoney])
                break
            elif((not staffNanme) and (not table.cell(row, 25).value) ):
                continue

            currentStaff+=1
            if(table.cell(row, 21).value):
                currentInsurance+=1
            bankMoney = table.cell(row, 23).value
            realMoney = table.cell(row, 24).value
            sumMoney = table.cell(row, 25).value

            if(bankMoney):
                personInBnak[staffNanme]=float(bankMoney)
            if(realMoney):
                personInReal[staffNanme]=float(realMoney)

    # add salary into result according to payment type
    result["departmentInfo"]=departmentInfo

    result["personInBnak"]=personInBnak

    result["personInReal"]=personInReal


def WriteXls(result):

    rb = xlrd.open_workbook("sheet/月度工资汇总.xls", formatting_info=True)
    staffList=GetStaffList()
    wb = copy(rb)
    table = wb.get_sheet(0)
    index=2

    # output staff salary in bank
    for staffInfo in  staffList:
        staff=staffInfo[0]
        XLFormat.setOutCell(table, 0, index, staffInfo[2])
        XLFormat.setOutCell(table, 1, index, staff)
        XLFormat.setOutCell(table, 2, index, staffInfo[1])
        '''
        if(result["staffList"][staff][1] != 0.0):
            XLFormat.setOutCell(table, 6, index, result["staffList"][staff][1]-result["personInBnak"][staff])
        else:
            XLFormat.setOutCell(table, 6, index, 0)
        result["staffList"].pop(staff)
        '''
        if(staff in  result["personInBnak"].keys()):
            XLFormat.setOutCell(table, 3, index, result["personInBnak"][staff])
            XLFormat.setOutCell(table, 5, index, result["personInBnak"][staff])

            result["personInBnak"].pop(staff)
        else:
            XLFormat.setOutCell(table, 6, index, "未找到该员工本月记录，请检查部门工资明细")
        index+=1

    for staff in  result["personInBnak"].keys():
        XLFormat.setOutCell(table, 1, index, staff)
        XLFormat.setOutCell(table, 3, index, result["personInBnak"][staff])
        XLFormat.setOutCell(table, 6, index, "未找到该员工信息，请检查StaffList.config文件")
        index+=1

    # output department Info
    index=5
    table1 = wb.get_sheet(1)
    for department in result["departmentInfo"]:
        for i in range(len(department)):
            XLFormat.setOutCell(table1, i, index, department[i])
        index+=1
    # output staff salary in real
    index=2
    table2 = wb.get_sheet(2)
    for staff in  result["personInReal"].keys():
        if(staff ):
            XLFormat.setOutCell(table2, 1, index, staff)
            XLFormat.setOutCell(table2, 3, index, result["personInReal"][staff])
            XLFormat.setOutCell(table2, 5, index, result["personInReal"][staff])
            index+=1

    wb.save('月度工资汇总.xls')

def GetStaffList():
    staffList = []
    tree = ET.parse("StaffList.config")     #打开xml文档
    root = tree.getroot()         #获得root节点

    for departmentNode in root:
        for staffNode in departmentNode:
            staffList.append([staffNode.get('Name'),staffNode.get('bankNo'),staffNode.get('No')])
    return staffList
    '''
    staffList = {}
    lastTem=[]
    table = data.sheets()[0]
    rows = table.nrows
    for row in range(2,rows):
        if(not table.cell(row, 1).value):
            if(not table.cell(row+1, 1)):
                break
            else:
                continue

        if(not table.cell(row,3).value):
            money=0.0
        else:
            money=table.cell(row,3).value
        staffList[table.cell(row, 1).value]=[table.cell(row,2).value,money]
        lastTem.append([table.cell(row, 1).value,table.cell(row,2).value])
    result["staffList"]=staffList
    result['lastTem']=lastTem
    '''

if __name__ == '__main__':
    result = {}
    #while not result:
    ReadXls(result)
    WriteXls(result)
    print("运行成功，结果保存在 月度工资汇总.xls中")
    os.system("pause")
