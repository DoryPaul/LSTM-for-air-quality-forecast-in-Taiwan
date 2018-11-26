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
    list = excel_data(r'C:\Users\文档\Desktop\1.xls')
    file = xlwt.Workbook()
    for i in range(len(list)):
        print('{}  & {} & {}  & {}  & {} & {} & {}  & {}  & {} &{} {}{}'.format(list[i][0],round(list[i][1],2),round(list[i][2],2),round(list[i][3],2),round(list[i][4],2),round(list[i][5],2),round(list[i][6],2),round(list[i][7],2),round(list[i][8],2),round(list[i][9],2),'\\','\\'))

