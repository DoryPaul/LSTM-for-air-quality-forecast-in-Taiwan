# -*- coding: utf-8 -*-

import xlrd
import xlwt
import random


def excel_data():
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(r'C:\Users\文档\Desktop\ORIGINAL.xls')
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

    fuhao=['#','*',' ','x','-','-']
    fuhao1 = ['警号', '星号', '空格', '埃克斯','负号','只有负号']
    tail=['\警号.xls','\星号.xls','\空格.xls','\埃克斯.xls','\负号.xls','\只有负号.xls']
    address=['宜兰','竹东','沙鹿','花莲','汐止','朴子','金门']
    for k in range(len(fuhao)):
        table = file.add_sheet(fuhao1[k])
        jinghao = 0
        if k!=2 and k!=5 and k!=4:
            for i in range(len(list)):
                for j in range(1,len(list[i])):
                    if fuhao[k] in str(list[i][j]) and '-'not in str(list[i][j]):
                        table.write(jinghao,0,list[i][0])
                        table.write(jinghao,1, address[j-1])
                        table.write(jinghao,2, list[i][j])
                #table.write(i,j,a)
                        jinghao+=1
            print(fuhao[k]+'总数：',jinghao)
                #table.write(i, j, list[i][j])

            file.save(r'C:\Users\文档\Desktop\temp1'+tail[k])
        if k==2:
            for i in range(len(list)):
                for j in range(1,len(list[i])):
                    if len(str(list[i][j]))==0:
                        table.write(jinghao,0,list[i][0])
                        table.write(jinghao,1, address[j-1])
                        table.write(jinghao,2, list[i][j])
                #table.write(i,j,a)
                        jinghao+=1
            print('空值总数：',jinghao)
                #table.write(i, j, list[i][j])

            file.save(r'C:\Users\文档\Desktop\temp1'+tail[k])
        if k==5:
            for i in range(len(list)):
                for j in range(1,len(list[i])):
                    if fuhao[k] in str(list[i][j]) and '#'not in str(list[i][j]) and '*'not in str(list[i][j]) and 'x'not in str(list[i][j]):
                        table.write(jinghao,0,list[i][0])
                        table.write(jinghao,1, address[j-1])
                        table.write(jinghao,2, list[i][j])
                #table.write(i,j,a)
                        jinghao+=1
            print('只有负号总数：',jinghao)
                #table.write(i, j, list[i][j])

            file.save(r'C:\Users\文档\Desktop\temp1'+tail[k])
        if k==4:
            for i in range(len(list)):
                for j in range(1,len(list[i])):
                    if fuhao[k] in str(list[i][j]):
                        table.write(jinghao,0,list[i][0])
                        table.write(jinghao,1, address[j-1])
                        table.write(jinghao,2, list[i][j])
                #table.write(i,j,a)
                        jinghao+=1
            print('混合总数：',jinghao)
                #table.write(i, j, list[i][j])

            file.save(r'C:\Users\文档\Desktop\temp1\混合.xls')

