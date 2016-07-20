# -*- encoding:UTF-8 -*-
import sitecustomize
import saleList
import StSalaryZLT2
import StSalaryDC2
import os


def GoFunction():
    print unicode('\n\n请根据需要的功能，选择对应的数字\n\n')
    print unicode('1.自络筒工人工资汇总\n2.挡车工人工资汇总\n3.销售情况汇总')
    num = raw_input()
    sheet = {}

    while num.isalpha() or num<'1' or num>'3':
        print unicode('输入错误，请重新输入','utf-8')
        num=raw_input()

    if num == '1':
        print unicode('自络筒工人工资汇总','utf-8')
        sheet = StSalaryZLT2.readXls()
        StSalaryZLT2.writeXls(sheet)
    elif num == '2':
        print unicode('挡车工人工资汇总','utf-8')
        sheet = StSalaryDC2.readXls()
        StSalaryDC2.writeXls(sheet)
    elif num == '3':
        print unicode('销售情况汇总','utf-8')
        sheet = saleList.readXls()
        saleList.writeXls(sheet)

if __name__ == '__main__':
    print unicode('/---------------南通双隆纺织科技有限公司--------------------/', 'utf-8')
    print unicode('/---------------企业管理软件v1.0---------------------------/', 'utf-8')
    GoFunction()
    print unicode('运行成功，结果保存在该目录的result.xls中','utf-8')
    os.system("pause")
