# -*- coding: utf-8 -*-

import xlrd
import xlwt
import random


def excel_data(i):
    try:
        # 打开Excel文件读取数据
        data = xlrd.open_workbook(r'C:\Users\文档\Desktop\修改后数据.xls')
        # 获取第一个工作表
        table = data.sheet_by_index(i)
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
    sheet=['PM2.5','PM10','CO','WIND_DIREC','WIND_SPEED','O3','NO2','RH','RAINFALL','AMB_TEMP','SO2']
    tail = ['\PM2.5.xls', '\PM10.xls', '\CO.xls', '\WIND_DIREC.xls', '\WIND_SPEED.xls', '\O3.xls', '\\NO2.xls', '\RH.xls', '\RAINFALL.xls','\AMB_TEMP.xls','\SO2.xls']
    for f in range(0,1):
        name1=sheet[f]
        name2=tail[f]
        list = excel_data(f)
        file = xlwt.Workbook()
        table = file.add_sheet(name1, cell_overwrite_ok=True)
        address = ['', '宜兰', '竹东', '沙鹿', '花莲', '汐止', '朴子', '金门']
        head=0
        tail=-1
        temp=[]
        record=[]
        for j in range(1,8):
            for i in range(len(list)-1):
                if list[i][j]!='' and list[i+1][j]!='':
                    if list[i][j]!='' and int(list[i+1][j])>300:
                        head=list[i][j]
                        temp.append((i,j))
                    if int(list[i][j])>300 and list[i+1][j]!='':
                        tail=list[i+1][j]
                    if int(list[i][j])<=300:
                        table.write(i,j,list[i][j])
                    if tail!=-1:
                #print('头为：',head)
                #print('尾为：',tail)
                        if head>=tail:
                            x=tail
                            y=head
                        else:
                            x=head
                            y=tail
                        for ind in temp:
                            a = round(random.uniform(x, y), 1)
                    #print(a)
                            record.append([list[ind[0]][0],address[ind[1]],list[i][j],a])
                            table.write(ind[0],ind[1],a)
                        head=0
                        tail=-1
                        temp=[]
                #print()

        file.save(r'C:\Users\文档\Desktop\temp2'+name2)

        file1 = xlwt.Workbook()
        table1 = file1.add_sheet(name1)
        for n in range(len(record)):
            for m in range(len(record[n])):
                table1.write(n,m,record[n][m])
        file1.save(r'C:\Users\文档\Desktop\temp3'+name2)

