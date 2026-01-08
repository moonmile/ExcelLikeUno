# コード補完チェック用

from excellikeuno import connect_calc
from excellikeuno.style.font import Font
from excellikeuno.style import BorderStyle
from excellikeuno.typing.calc import BorderLineStyle

(desktop, doc, sheet) = connect_calc() 
cell = sheet.cell(0, 0)  # A1 セルを取得
cell.text = "Hello, World!"  # 値を設定
cell.font.bold = True
cell.font.size = 16

range = sheet.range("A1:C1")
range.font = Font(name="Arial", size=16, color=0xFF0000)

text_shape = sheet.shapes.add_text_shape(1000,1000,4000,500,text="Sample")
text_shape.font.size = 14
text_shape.font.name = "Courier New"


cell.borders.top.color = 0x0000FF  # 青色の上罫線を設定
cell.borders.top.weight = 50  # 線の太さを設定
cell.borders.top.line_style = BorderLineStyle.SOLID  # 実線 


# フォント
font = Font(name="Times New Roman", size=12, color=0x000000)

# 罫線
sheet.cell(0, 0).borders.all = BorderStyle(
    color=0x00FF00,  # 緑色
    weight=30,       # 線の太さ
    line_style=BorderLineStyle.DASHED  # 破線
)
sheet.range("B2:D5").borders.all = BorderStyle(
    color=0xFF00FF,  # 紫色
    weight=20,       # 線の太さ
    line_style=BorderLineStyle.SOLID  # 実線
)






