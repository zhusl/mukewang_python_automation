# Python办公自动化
## [在线视频](https://www.bilibili.com/video/BV11p4y1n7Dx/)
> [第1章 导学]()<br>
> [第2章 环境的搭建](https://www.bilibili.com/video/BV1NV411q7gV)<br>
> [第3章 xlrd+xlwt读取/写入Excel数据](https://www.bilibili.com/video/BV1BN411R7Q4)<br>
> [第4章 xlsxwriter生成图表 ](https://www.bilibili.com/video/BV1jr4y1N7Ty)<br>
> [第5章 玩转Word自动化 ](https://www.bil<br>ibili.com/video/BV1fA411T7oA)<br>
> [第6章 玩转PPT自动化](https://www.bilibili.com/video/BV1Kr4y1N7md)<br>

## 依赖
[xlwt](https://pypi.org/project/xlwt/1.0.0/)<br>
[xlrd](https://pypi.org/project/xlrd/)<br>

如果已经安装pip的话直接运行
```bash
pip install xlwt
pip install xlrd
```
###### 注意
xlwt只支持2003-2013的```xls```格式，不支持2007版以后的```xlsx```格式

## 数据
|A	|B    |C	 |D	 |E    |F	  |G|
|  :-:  |  :-:  | :-:   | :-:  |  :-:   |:-:    |:-:|
| | | |1月份销售明细表 |	|	|				
|日期	|货号    |颜色	 |尺码	 |原价    |折扣	  |备注|
|1.20	|X0001	|红	    |M	    |199	|80%	|无|
|1.20	|X0002	|绿	    |L	    |199	|70%	|无|
|1.25   |X0003	|蓝	    |XL	    |200	|90%	|无|
|1.28	|X0004	|黑	    |XXL	|399	|95%	|无|
|1.29	|X0005	|灰	    |S	    |80	    |50%	|特价商品|
|1.30	|X0006	|白	    |XS	    |500	|90%	|热销新品|
|1.30	|X0007	|黑	    |M	    |199	|90%	|无|
|1.31	|X0008	|红	    |XL	    |210	|90%	|无|
|1.31	|X0009	|灰	    |L	    |50	    |90%	|无|


### 格式化
```D:\python3\Lib\site-packages\xlwt```中，因为的的python安装目录是```D:\python3```
##### 样式Style.py
```python
# Text values for colour indices. "grey" is a synonym of "gray".
# The names are those given by Microsoft Excel 2003 to the colours
# in the default palette. There is no great correspondence with
# any W3C name-to-RGB mapping.
_colour_map_text = 
aqua 0x31
black 0x08
blue 0x0C
blue_gray 0x36
bright_green 0x0B
brown 0x3C
coral 0x1D
cyan_ega 0x0F
dark_blue 0x12
dark_blue_ega 0x12
dark_green 0x3A
dark_green_ega 0x11
dark_purple 0x1C
dark_red 0x10
dark_red_ega 0x10
dark_teal 0x38
dark_yellow 0x13
gold 0x33
gray_ega 0x17
gray25 0x16
gray40 0x37
gray50 0x17
gray80 0x3F
green 0x11
ice_blue 0x1F
indigo 0x3E
ivory 0x1A
lavender 0x2E
light_blue 0x30
light_green 0x2A
light_orange 0x34
light_turquoise 0x29
light_yellow 0x2B
lime 0x32
magenta_ega 0x0E
ocean_blue 0x1E
olive_ega 0x13
olive_green 0x3B
orange 0x35
pale_blue 0x2C
periwinkle 0x18
pink 0x0E
plum 0x3D
purple_ega 0x14
red 0x0A
rose 0x2D
sea_green 0x39
silver_ega 0x16
sky_blue 0x28
tan 0x2F
teal 0x15
teal_ega 0x15
turquoise 0x0F
violet 0x14
white 0x09
yellow 0x0D
```
##### 对齐Formatting.py
在``` D:\python3\Lib\site-packages\xlwt\Formatting.py```
```py
    def __init__(self):
        self.horz = self.HORZ_GENERAL
        self.vert = self.VERT_BOTTOM
        self.dire = self.DIRECTION_GENERAL
        self.orie = self.ORIENTATION_NOT_ROTATED
        self.rota = self.ROTATION_0_ANGLE
        self.wrap = self.NOT_WRAP_AT_RIGHT
        self.shri = self.NOT_SHRINK_TO_FIT
        self.inde = 0
        self.merg = 0

    class Alignment(object):
    # 水平方向
    HORZ_GENERAL                = 0x00
    HORZ_LEFT                   = 0x01
    HORZ_CENTER                 = 0x02
    HORZ_RIGHT                  = 0x03
    HORZ_FILLED                 = 0x04
    HORZ_JUSTIFIED              = 0x05 # BIFF4-BIFF8X
    HORZ_CENTER_ACROSS_SEL      = 0x06 # Centred across selection (BIFF4-BIFF8X)
    HORZ_DISTRIBUTED            = 0x07 # Distributed (BIFF8X)

    # 垂直方向
    VERT_TOP                    = 0x00
    VERT_CENTER                 = 0x01
    VERT_BOTTOM                 = 0x02
    VERT_JUSTIFIED              = 0x03 # Justified (BIFF5-BIFF8X)
    VERT_DISTRIBUTED            = 0x04 # Distributed (BIFF8X)

    DIRECTION_GENERAL           = 0x00 # BIFF8X
    DIRECTION_LR                = 0x01
    DIRECTION_RL                = 0x02

    ORIENTATION_NOT_ROTATED     = 0x00
    ORIENTATION_STACKED         = 0x01
    ORIENTATION_90_CC           = 0x02
    ORIENTATION_90_CW           = 0x03

    ROTATION_0_ANGLE            = 0x00
    ROTATION_STACKED            = 0xFF

    WRAP_AT_RIGHT               = 0x01
    NOT_WRAP_AT_RIGHT           = 0x00

    SHRINK_TO_FIT               = 0x01
    NOT_SHRINK_TO_FIT           = 0x00    
```

