'''
xlwt写入Excel步骤
    1.创建工作簿
    2.创建工作表
    3.填充工作表内容
    4.保存
'''
import xlwt
# 1.创
wb=xlwt.Workbook()
# 2.创建工作表
ws=wb.add_sheet("CNY")
ws2=wb.add_sheet("Image") # 创建第2个工作表

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

########################################
### 第2个工作表写入
#ws2.insert_bitmap("2017年货币兑换表.png",0,0) # 会报错
ws2.insert_bitmap("2017.bmp",0,0)
'''插入图片
    insert_bitmap(self, filename, row, col, x=0, y=0, scale_x=1, scale_y=1)
    insert_bitmap_data(self, data, row, col, x=0, y=0, scale_x=1, scale_y=1
    必须是 .bmp (不能仅仅改文件名的那种):https://jingyan.baidu.com/article/72ee561a801002a06138df83.html
'''

# 4.保存
wb.save("2019-CNY.xls") #wlxt只支持xls





