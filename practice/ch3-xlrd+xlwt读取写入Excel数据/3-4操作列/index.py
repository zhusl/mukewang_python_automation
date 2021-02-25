#操作excel列
import xlrd

# 打开excel文件
data=xlrd.open_workbook("data1.xlsx")
sheet1=data.sheet_by_index(0) #根据索引 获取第一个工作表

# 获取当前Sheet下的列数
print('{}： "{}"'.format('第1个Sheet的列数',sheet1.ncols))

# 获取行的内容
print('{}： "{}"'.format('第1个Sheet的第1列内容',sheet1.col(0))) # col返回是单元格对象组成的列表
print('{}： "{}"'.format('第1个Sheet的第2列内容',sheet1.col(1)))

# 获取一列单元格的数据类型
print(sheet1.col_types(1))
print(sheet1.col_types(2))

# 获取具体单元格的数据对象(先从列开始)
print(sheet1.col(2)[1])  # 获取第2行第3列的数据对象

# 获取具体单元格的数据的值(先从列开始)
print(sheet1.col(2)[1].value)  # 获取第2行第3列的数据的值

# 获取列的数据的值
print(sheet1.col_values(1)) # 获取第2列的数据的值
print(sheet1.col_values(1)[3]) # 获取第4行第2列的数据的值

# 没有获取一列元素的长度(就是有几行)的方法
# print('{}： "{}"'.format('第2列长度',sheet1.col_len(1)))

