import xlrd

# 打开excel文件
data=xlrd.open_workbook("data1.xlsx")
"""
data.sheet_loaded(0)  # 加载第1个工作表
data.unload_sheet(0)  # 卸载第1个工作表
"""
# 打印工作表
print(data.sheets())
print(data.sheets()[0]) #第1个
print(data.sheets()[1]) #第2个
 
# 打印工作表是否加载
print(data.sheet_loaded(0))
print(data.sheet_loaded(1))

# 获取所有工作表的名字
print('{}： "{}"'.format('所有工作表的名字:', data.sheet_names()))
# 获取所有工作表的数量——用于索引
print('{}： "{}"'.format('所有工作表的数量:', data.nsheets))
# 获取工作表
print(data.sheet_by_index(0)) #根据索引 获取第一个工作表
print(data.sheet_by_name("Sheet1")) #根据名称 获取第一个工作表

   