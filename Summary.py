import xlrd
import glob
import xlwt
from xlutils.copy import copy
import XLFormat


def readXls():
    print('\n------------------------------------')
    print('公司工资统计/ v1.2')
    print('请将需要统计的报表放在该目录的‘销售数据’文件夹下,并从0.xls开始顺序命名')
    print('请输入待汇总的表格份数')

    result = {}
    personInBnak = {}
    personInReal = {}
    departmentList = {}
    departmentInfo = {}
    file =glob.glob('data/*.xls')
    print(file)
    try:
        data = xlrd.open_workbook(file[0])
    except:
        print('没有找到待读取的%s，操作失败！请检查您的文件\n\n' %file[0])
        return

    print('%s读取成功，正在处理......' %file[0])
    workerList = {}
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
        workerList[table.cell(row, 1).value]=[table.cell(row,2).value,money]
    result["workerList"]=workerList


    for sheet in range(3,10):
        table = data.sheets()[sheet]
        rows = table.nrows
        for row in range(4, rows):
            workerNanme = table.cell(row, 1).value
            if( (not table.cell(row, 25).value) and (not workerNanme) and (not table.cell(row+1, 1).value) ):
                break

            bankMoney = table.cell(row, 23).value
            realMoney = table.cell(row, 24).value
            sumMoney = table.cell(row, 25).value

            if(bankMoney):
                personInBnak[workerNanme]=float(bankMoney)
            if(realMoney):
                personInReal[workerNanme]=float(realMoney)

    if("personInBnak" in result):
        result["personInBnak"].update(personInBnak)
    else:
        result["personInBnak"]=personInBnak

    if("personInReal" in result):
        result["personInReal"].update(personInReal)
    else:
        result["personInReal"]=personInReal
    return result


def writeXls(result):

    rb = xlrd.open_workbook("工资.xls", formatting_info=True)

    wb = copy(rb)
    table = wb.get_sheet(0)
    index=2
    for worker in  result["personInBnak"].keys():
        if(worker ):
            XLFormat.setOutCell(table, 1, index, worker)
            if(worker in result["workerList"] ):
                XLFormat.setOutCell(table, 2, index, result["workerList"][worker][0])
                if(result["workerList"][worker][1] != 0.0):
                    XLFormat.setOutCell(table, 6, index, result["workerList"][worker][1]-result["personInBnak"][worker])
                else:
                    XLFormat.setOutCell(table, 6, index, 0)
                result["workerList"].pop(worker)
            XLFormat.setOutCell(table, 3, index, result["personInBnak"][worker])
            XLFormat.setOutCell(table, 5, index, result["personInBnak"][worker])
            index+=1
        #XLFormat.setOutCell(table, 6, i, sheet[key])

    for worker in  result["workerList"].keys():
        XLFormat.setOutCell(table, 1, index, worker)
        index+=1
    wb.save('工资1.xls')


if __name__ == '__main__':
    result = {}
    #while not result:
    result = readXls()
    writeXls(result)
    print("运行成功，结果保存在该目录的 销售报表.xls中")
