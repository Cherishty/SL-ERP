import xlrd
import xlwt
from xlutils.copy import copy
import XLFormat


def readXls():
    print('\n------------------------------------')
    print('挡车工序 工资计算v1.4')
    print('请将需要统计的报表放在该目录的data文件夹下,并从0.xls开始顺序命名')
    print('请输入待汇总的表格份数')

    try:
        rb = xlrd.open_workbook('销售报表.xls')
    except:
        print('没有找到 工资总表.xls，操作失败！请检查您的路径或命名\n\n')
        return

    sheet = {}
    num = input()
    while num.isalpha():
        print('输入错误，请重新输入合法数字')
        num = input()

    for each in range(int(num)):
        try:
            data = xlrd.open_workbook('dataDC/' + str(each) + '.xls')
        except:
            print('没有找到待读取的%d.xls，操作失败！请检查您的路径或命名\n\n' % each)
            print('------------------------------------\n\n')
            return
        else:
            print('%d.xls读取成功，正在处理......' % each)
            table = data.sheets()[0]
            list = []
            cell_B1 = table.cell(0, 9).value
            list.append(cell_B1)
            morning = table.cell(35, 13).value
            noon = table.cell(35, 14).value
            night = table.cell(35, 15).value
            day = table.cell(35, 16).value + table.cell(36, 16).value
            twelve = table.cell(36, 16).value
            bus = table.cell(35, 17).value
            ill = table.cell(35, 18).value
            kuang = table.cell(35, 19).value
            list.append(morning)
            list.append(noon)
            list.append(night)
            list.append(day)
            list.append(twelve)
            list.append(bus)
            list.append(ill)
            list.append(kuang)
            base = table.cell(31, 22).value
            super = table.cell(31, 23).value
            list.append(base)
            list.append(super)
            sheet[list[0]] = list
    return sheet


def writeXls(sheet):
    rb = xlrd.open_workbook('工资总表.xls', formatting_info=True)
    wb = copy(rb)
    table = wb.get_sheet(0)
    i = 3
    for list in sheet.values():
        j = 1
        for each in list:
            XLFormat.setOutCell(table, j, i, each)
            XLFormat.setOutCell(table, 18, i, xlwt.Formula('SUM(K%d:Q%d)-R%d' % (i + 1, i + 1, i + 1)))
            XLFormat.setOutCell(table, 23, i, xlwt.Formula('S%d-T%d-U%d' % (i + 1, i + 1, i + 1)))
            j += 1
        i += 1
    for each in range(i, 41):
        XLFormat.setOutCell(table, 18, each, xlwt.Formula('SUM(K%d:Q%d)-R%d' % (i + 1, i + 1, i + 1)))
        XLFormat.setOutCell(table, 23, each, xlwt.Formula('S%d-T%d-U%d' % (i + 1, i + 1, i + 1)))
    for j in range(75, 89):
        XLFormat.setOutCell(table, j - ord('A'), 41, xlwt.Formula('SUM(%c4:%c41)' % (chr(j), chr(j))))
    wb.save('工资总表.xls')


if __name__ == '__main__':
    sheet = {}
    while not sheet:
        sheet = readXls()
    writeXls(sheet)
    print('运行成功，结果保存在该目录的 工资总表.xls中')
