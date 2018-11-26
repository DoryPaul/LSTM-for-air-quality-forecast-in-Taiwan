import xlrd
import xlwt

file = xlwt.Workbook()
table = file.add_sheet('宜兰1.xls')
n=0
for i in range(1,1826):
    for j in range(0,24):
        table.write(n, 0, j)
        #num=i%365
       # if num!=0:
        #    table.write(n, 0, num)
       # else:
         #   table.write(n, 0, 365)
        n+=1
file.save(r'C:\Users\文档\Desktop\宜兰1.xls')