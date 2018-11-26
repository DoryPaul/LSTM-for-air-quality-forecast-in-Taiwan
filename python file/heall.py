# -*- coding: utf-8 -*-

import xlrd
import xlwt


def excel_data(name):
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(name)
        # 获取第一个工作表
        table = data.sheet_by_index(9)
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
    list = excel_data(r'C:\Users\文档\Desktop\宜兰\宜兰合.xls')
    list1 = excel_data(r'C:\Users\文档\Desktop\竹东\竹东合.xls')
    list2 = excel_data(r'C:\Users\文档\Desktop\沙鹿\沙鹿合.xls')
    list3 = excel_data(r'C:\Users\文档\Desktop\花莲\花莲合.xls')
    list4 = excel_data(r'C:\Users\文档\Desktop\汐止\汐止合.xls')
    list5 = excel_data(r'C:\Users\文档\Desktop\朴子\朴子合.xls')
    list6 = excel_data(r'C:\Users\文档\Desktop\金门\金门合.xls')
    file = xlwt.Workbook()
    table = file.add_sheet('宜兰1.xls')
    for i in range(len(list)):
        table.write(i,0,list[i][0])
        table.write(i,1,list[i][1])
    for i in range(len(list1)):
        table.write(i,2,list1[i][1])
    for i in range(len(list2)):
        table.write(i, 3, list2[i][1])
    for i in range(len(list3)):
        table.write(i, 4, list3[i][1])
    for i in range(len(list4)):
        table.write(i, 5, list4[i][1])
    for i in range(len(list5)):
        table.write(i, 6, list5[i][1])
    for i in range(len(list6)):
        table.write(i, 7, list6[i][1])
    file.save(r'C:\Users\文档\Desktop\宜兰1.xls')

