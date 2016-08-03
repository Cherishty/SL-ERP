import xlrd
import xlwt
from xlutils.copy import copy
import XLFormat


def readXls():
    print('\n------------------------------------')
    print('销售汇总 v1.2')
    print('请将需要统计的报表放在该目录的‘销售数据’文件夹下,并从0.xls开始顺序命名')
    print('请输入待汇总的表格份数')

    companyList = {}
    num = input()
    while num.isalpha():
        print('输入错误，请重新输入合法数字')
        num = input()

    for each in range(int(num)):
        try:
            data = xlrd.open_workbook('销售数据/' + str(each) + '.xls')
        except:
            print('没有找到待读取的%d.xls，操作失败！请检查您的路径或命名\n\n' % each)
            return

        print('%d.xls读取成功，正在处理......' % each)
        table = data.sheets()[0]
        row = table.nrows
        for eachRow in range(2, row):
            companyName = table.cell(eachRow, 2).value
            if not companyName:
                break
            money = table.cell(eachRow, 6).value
            if companyName in companyList:
                companyList[companyName] = companyList[companyName] + float(money)
            else:
                companyList[companyName] = float(money)
    return companyList


def writeXls(sheet):
    rb = xlrd.open_workbook('sheet/销售报表.xls', formatting_info=True)
    wb = copy(rb)
    table = wb.get_sheet(0)
    i = 2
    for key in sheet:
        XLFormat.setOutCell(table, 2, i, key)
        XLFormat.setOutCell(table, 6, i, sheet[key])
        i += 1
    wb.save('销售报表.xls')


if __name__ == '__main__':
    sheet = {}
    while not sheet:
        sheet = readXls()
    writeXls(sheet)
    print("运行成功，结果保存在该目录的 销售报表.xls中")
