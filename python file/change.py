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
        for row in range(0, nrows):
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
    list = excel_data(r'C:\Users\文档\Desktop\朴子\朴子合.xls')

    file = xlwt.Workbook()
    #sheet=['SO2']
    #tail = ['\SO2.xls']
    sheet=['PM2.5','PM10','CO','WIND_DIREC','WIND_SPEED','O3','NO2','RH','RAINFALL','AMB_TEMP','SO2']
    tail = ['\PM2.5.xls', '\PM10.xls', '\CO.xls', '\WIND_DIREC.xls', '\WIND_SPEED.xls', '\O3.xls', '\\NO2.xls', '\RH.xls', '\RAINFALL.xls','\AMB_TEMP.xls','\SO2.xls']

    for k in range(len(sheet)):
        table = file.add_sheet(sheet[k])
        a=0
        dict=['-0','-1','-2','-3','-4','-5','-6','-7','-8','-9','-10','-11','-12','-13','-14','-15','-16','-17','-18','-19','-20','-21','-22','-23']
        for i in list:
            if i[2]==sheet[k]:
                for j in range(3,len(i)):
                #print(i[0]+' 1')
                    x=a%24
                    table.write(a,0,i[0]+dict[x])
                    table.write(a,1,i[j])
                    a+=1

        file.save(r'C:\Users\文档\Desktop\temp'+tail[k])

