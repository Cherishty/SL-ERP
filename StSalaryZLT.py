import xlrd
import xlwt
from xlutils.copy import copy
import XLFormat


def readXls():
    print('\n------------------------------------')
    print('自络筒工序 工资计算v1.4')
    print('请将需要统计的报表放在该目录的data文件夹下,并从0.xls开始顺序命名')
    print('请输入待汇总的表格份数')
    sheet = {}

    num = input()
    while num.isalpha():
        print('输入错误，请重新输入合法数字')
        num = input()

    for each in range(int(num)):
        try:
            data = xlrd.open_workbook('dataZLT/' + str(each) + '.xls')
        except:
            print('没有找到该目录或该文件，请检查您的路径或命名\n\n')
            print('------------------------------------\n\n\n')
            return
        else:
            table = data.sheets()[0]
            list = []
            name = table.cell(0, 4).value
            list.append(name)
            morning = table.cell(34, 5).value
            noon = table.cell(34, 6).value
            night = table.cell(34, 7).value
            day = table.cell(34, 8).value
            twelve = table.cell(35, 8).value
            bus = table.cell(34, 9).value
            ill = table.cell(34, 10).value
            kuang = table.cell(34, 11).value
            list.append(morning)
            list.append(noon)
            list.append(night)
            list.append(day)
            list.append(twelve)
            list.append(bus)
            list.append(ill)
            list.append(kuang)
            base = table.cell(12, 18).value
            super = table.cell(12, 19).value
            list.append(base)
            list.append(super)
            sheet[list[0]] = list

    return sheet


def writeXls(sheet):
    rb = xlrd.open_workbook('result.xls', formatting_info=True)
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
    wb.save('result.xls')


if __name__ == '__main__':
    sheet = {}
    while not sheet:
        sheet = readXls()

    writeXls(sheet)
    print("run successfully!")
