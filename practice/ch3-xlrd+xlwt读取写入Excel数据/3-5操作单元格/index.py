#操作excel单元格
import xlrd

# 打开excel文件
data=xlrd.open_workbook("data1.xlsx")
sheet1=data.sheet_by_index(0) #根据索引 获取第一个工作表

# 获取某一单元格内容
print(sheet1.cell(1,2)) #cell返回的是一个单元格对象

# 获取某一单元格的数据类型
print(sheet1.cell(1,2).ctype) # 使用.ctype来获取
print(sheet1.cell_type(1,2)) # 使用专门的函数

# 获取具体单元格的数据的值
print(sheet1.cell(1,2).value) # 使用.value来获取
print(sheet1.cell_value(1,2))  # 使用专门的函数

