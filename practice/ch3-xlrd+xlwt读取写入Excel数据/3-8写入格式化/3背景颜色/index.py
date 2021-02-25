'''
单元格格式设置，样式设置
'''
import xlwt
# 1.创
wb=xlwt.Workbook()
# 2.创建工作表
ws=wb.add_sheet("CNY")
ws2=wb.add_sheet("Image") # 创建第2个工作表

# 插入样式
## 参考："D:\python3\Lib\site-packages\xlwt\Style.py"
### 实例化一个字体样式
titlestyle=xlwt.XFStyle() #初始化样式
titlefont=xlwt.Font()
titlefont.name="宋体"
titlefont.bold=True #加粗
titlefont.height=11*20 #11是字号 20是一个衡量单位
titlefont.colour_index=0x08 # 字体设置为黑色

titlestyle.font=titlefont # 添加字体

### 设置边框
borders=xlwt.Borders()
borders.right=xlwt.Borders.DASHED #右侧是虚线
borders.bottom=xlwt.Borders.DOTTED #地下是点线

titlestyle.borders=borders #添加边框

### 背景颜色
datestyle = xlwt.XFStyle()
bgcolor = xlwt.Pattern()
bgcolor.pattern=xlwt.Pattern.SOLID_PATTERN
bgcolor.pattern_fore_colour=22 #灰色
datestyle.pattern=bgcolor

### 对齐样式
##### 参考：D:\python3\Lib\site-packages\xlwt\Formatting.py
### 实例化一个单元格对其方式
cellalign=xlwt.Alignment()
cellalign.horz=0x02 #居中对齐
# 也可以 cellalign.horz=0x02 #居中对齐
cellalign.vert=0x01

titlestyle.alignment=cellalign #设置对齐方式


# 3.填充工作表内容
'''合并单元格
  >>> help (xlwt.Worksheet)
      write_merge(self, r1, r2, c1, c2, label='', style=<xlwt.Style.XFStyle object at 0x0000018769444E50>)
'''      
## 默认风格
# ws.write_merge(0,1,0,5, "2019年货币兑换表") #第1-2行合并，第1-6列合并
ws.write_merge(0,1,0,5, "2019年货币兑换表", titlestyle) #自定义风格
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
        if j==0:
            ws.write(i+2, j, val, datestyle)   #自定义背景颜色
        else:                                  # 默认背景颜色
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





