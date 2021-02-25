'''
xlwt写入Excel步骤
    1.创建工作簿
    2.创建工作表
    3.填充工作表内容
    4.保存
'''
import xlwt
# 1.创建工作簿
wb=xlwt.Workbook()
# 2.创建工作表
ws=wb.add_sheet("CNY")
# 3.填充工作表内容
'''合并单元格
  >>> help (xlwt.Worksheet)
      write_merge(self, r1, r2, c1, c2, label='', style=<xlwt.Style.XFStyle object at 0x0000018769444E50>)
'''      
ws.write_merge(0,1,0,5, "2019年货币兑换表") #第1-2行合并，第1-6列合并

'''写入数据
    >>> help (xlwt.Worksheet)
        write(self, r, c, label='', style=<xlwt.Style.XFStyle object at 0x0000018769444E50>)
'''
data=(   # 需要写入的数据
    ("Data", "英镑", "人民币", "港币" , "日元", "美元"),  #第一行数据
    ("01/01/2019", 8.722551, 1, 0.877885, 0.062722, 6.8759),   #第二行数据
    ("02/01/2019", 8.634922, 1, 0.875731, 0.062773, 6.8601))   #第三行数据

for i, item in enumerate(data):            #每一行数据
    for j, val in enumerate(item):         #每一行的每一列数据
        ws.write(i+2, j, val)              # 因为前2行有数据了，所以需要跳过

# 4.保存
wb.save("2019-CNY.xls") #wlxt只支持xls





