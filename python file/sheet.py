# -*- coding: utf-8 -*-

import xlrd
import xlwt


def excel_data():
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(r'C:\Users\文档\Desktop\宜兰\宜兰合.xls')
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
                if col!=1 and col!=2:
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
    dict=['-0','-1','-2','-3','-4','-5','-6','-7','-8','-9','-10','-11','-12','-13','-14','-15','-16','-17','-18','-19','-20','-21','-22','-23']
    for i in list:
        for j in range(1,len(i)):
            print(i[0]+' 1')
            x=a%24
            table.write(a,0,i[0]+dict[x])
            table.write(a,1,i[j])
            a+=1
    file.save(r'C:\Users\文档\Desktop\宜兰1.xls')

