# -*- coding: utf-8 -*-

import xlrd
import xlwt
import random


def excel_data():
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(r'C:\Users\文档\Desktop\missing data\修改后数据.xls')
        # 获取第一个工作表
        table = data.sheet_by_index(10)
        # 获取行数
        nrows = table.nrows
        # 获取列数
        ncols = table.ncols
        # 定义excel_list
        excel_list = []
        total=[]
        for row in range(nrows):
            for col in range(ncols):
                # 获取单元格数据
                cell_value = table.cell(row, col).value
                # 把数据追加到excel_list中
                excel_list.append(cell_value)
            total.append(excel_list)
            excel_list=[]
        return total
    except:
        print('wrong')


if __name__ == "__main__":
    list = excel_data()
    file = xlwt.Workbook()

    table = file.add_sheet('改')
    jinghao = 0

    for i in range(len(list)):
        for j in range(1,len(list[i])):
            if '#' in str(list[i][j]) or '*' in str(list[i][j]) or 'x' in str(list[i][j]):
                table.write(i,j,list[i][j][:-1])
                jinghao+=1
            else:
                table.write(i, j, list[i][j])

                #table.write(i,j,a)
            file.save(r'C:\Users\文档\Desktop\temp1\temp.xls')
    print('总修改数：',jinghao)