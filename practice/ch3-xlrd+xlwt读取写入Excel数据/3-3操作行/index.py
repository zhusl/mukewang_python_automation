#操作excel行
import xlrd
# 打开excel文件
data=xlrd.open_workbook("data1.xlsx")
sheet1=data.sheet_by_index(0) #根据索引 获取第一个工作表

# 获取当前Sheet下的行数
print('{}： "{}"'.format('第1个Sheet的行数',sheet1.nrows))
sheet2=data.sheet_by_index(1) 
print('{}： "{}"'.format('第2个Sheet的行数',sheet2.nrows))

# 获取行的内容
print('{}： "{}"'.format('第1个Sheet的第1行内容',sheet1.row(0))) # row返回是单元格对象组成的列表
print('{}： "{}"'.format('第1个Sheet的第2行内容',sheet1.row(1))) 

# 获取一行单元格的数据类型
print(sheet1.row_types(1))
print(sheet1.row_types(2))
'''
SheetObject.row_types(rowx[, start_colx=0, end_colx=None])
获取sheet中第rowx+1行从start_colx列到end_colx列的单元类型，返回值为array.array类型。
单元类型ctype：empty为0，string为1，number为2，date为3，boolean为4， error为5（左边为类型，右边为类型对应的值）；
''' 

# 获取具体单元格的数据对象(先从行开始)
print(sheet1.row(1)[2])  # 获取第2行第3列的数据对象

# 获取具体单元格的数据的值(先从行开始)
print(sheet1.row(1)[2].value)  # 获取第2行第3列的数据的值

# 获取行的数据的值
print(sheet1.row_values(1)) # 获取第2行的数据的值
print(sheet1.row_values(1)[3]) # 获取第2行第4列的数据的值

# 获取一行元素的长度(就是有几列)
print('{}： "{}"'.format('第2行长度',sheet1.row_len(1)))
