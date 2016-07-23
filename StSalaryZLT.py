# -*- encoding:UTF-8 -*-
import sitecustomize
import xlrd
import xlwt


def readXls():
    print('------------------------------------')
    print unicode('自络筒工序 工资计算v1.3', 'utf-8')  # decode via unicode function
    print unicode('请将需要统计的报表放在该目录的data文件夹下,并从0.xls开始顺序命名', 'utf-8')
    print unicode('请输入待汇总的表格份数', 'utf-8')
    sheet = {}

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
    file = xlwt.Workbook(encoding='utf-8')  # encoding='utf-8' to support chinese output
    table = file.add_sheet('1')
    i = 0
    k = 0
    title = ['姓名', '早', '中', '夜', '天数', '12h', '事假', '病假', '旷工', '基本工资', '超产工资']
    for each in title:
        table.write(i, k, each)
        k += 1
    i += 1
    for list in sheet.values():
        j = 0
        for each in list:
            table.write(i, j, each)
            j += 1
        i += 1
    file.save('result.xls')


if __name__ == '__main__':
    sheet = {}
    while not sheet:
        sheet = readXls()

    writeXls(sheet)
    print "run successfully!"
