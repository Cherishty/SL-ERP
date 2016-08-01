import SalesList
import StSalaryZLT
import StSalaryDC
import os


def GoFunction():
    print('\n\n请根据需要的功能，选择对应的数字\n\n')
    print('1.自络筒工人工资汇总\n2.挡车工人工资汇总\n3.销售情况汇总')
    sheet = {}

    num = input()
    while num.isalpha() or int(num) < 1 or int(num) > 3:
        print('输入错误，请重新输入合法数字')
        num = input()

    if num == '1':
        while not sheet:
            sheet = StSalaryZLT.readXls()
        StSalaryZLT.writeXls(sheet)
    elif num == '2':
        while not sheet:
            sheet = StSalaryDC.readXls()
        StSalaryDC.writeXls(sheet)
    elif num == '3':
        while not sheet:
            sheet = SalesList.readXls()
        SalesList.writeXls(sheet)


if __name__ == '__main__':
    print('/---------------南通双隆纺织科技有限公司--------------------/')
    print('/---------------企业管理软件v1.2---------------------------/')
    GoFunction()
    print('运行成功，结果保存在该目录的相应表单中')
    os.system("pause")
