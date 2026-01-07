# コード補完チェック用

from excellikeuno import connect_calc
from excellikeuno.style.font import Font

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

text_shape.ha
