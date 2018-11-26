# -*- coding: utf-8 -*-

import xlrd
import xlwt


def excel_data(name):
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(name)
        # 获取第一个工作表
        table = data.sheet_by_index(0)
        # 获取行数
        nrows = table.nrows
        # 获取列数
        ncols = table.ncols
        # 定义excel_list
        excel_list = []
        total=[]
        for row in range(1, nrows):
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
    title1='金门'
    title2='金门'
    list = excel_data(r'C:\Users\文档\Desktop\金门\2013金門.xls')
    list1 = excel_data(r'C:\Users\文档\Desktop\金门\2014金門.xls')
    list2 = excel_data(r'C:\Users\文档\Desktop\金门\2015金門.xls')
    list3 = excel_data(r'C:\Users\文档\Desktop\金门\2016金門.xls')
    list4 = excel_data(r'C:\Users\文档\Desktop\金门\2017金門.xls')
    file = xlwt.Workbook()
    table = file.add_sheet('宜兰1.xls')

    for i in range(len(list)):
        for j in range(len(list[i])):
            #print(i[0]+' 1')
            table.write(i,j,list[i][j])

    for i in range(len(list1)):
        for j in range(len(list1[i])):
            #print(i[0]+' 1')
            table.write(i+len(list)+1,j,list1[i][j])

    for i in range(len(list2)):
        for j in range(len(list2[i])):
            #print(i[0]+' 1')
            table.write(i+len(list)+len(list1)+1,j,list2[i][j])

    for i in range(len(list3)):
        for j in range(len(list3[i])):
            #print(i[0]+' 1')
            table.write(i+len(list)+len(list1)+len(list2)+1,j,list3[i][j])

    for i in range(len(list4)):
        for j in range(len(list4[i])):
            #print(i[0]+' 1')
            table.write(i+len(list)+len(list1)+len(list2)+len(list3)+1,j,list4[i][j])

    file.save(r'C:\Users\文档\Desktop\宜兰1.xls')

