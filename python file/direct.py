# -*- coding: utf-8 -*-

import xlrd
import xlwt


def excel_data():
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(r'E:\hualian.xls')
        # 获取第一个工作表
        table = data.sheet_by_index(0)
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
    table = file.add_sheet('宜兰1.xls')
    a=0
    c=0
    b=''
    for i in list:
        a=i[3]
        if 337.5 < a and a <= 22.5:
            b='N'
        elif 22.5 < a and a <= 67.5:
            b='NE'
        elif 67.5 < a and a <= 112.5:
            b = 'E'
        elif 112.5 < a and a <= 157.5:
            b = 'SE'
        elif 157.5 < a and a <= 202.5:
            b = 'S'
        elif 202.5 < a and a <= 247.5:
            b = 'SW'
        elif 247.5 < a and a <= 292.5:
            b = 'W'
        elif 292.5 < a and a <= 337.5:
            b = 'NW'
        table.write(c,0,b)
        c+=1
    file.save(r'C:\Users\文档\Desktop\宜兰1.xls')

