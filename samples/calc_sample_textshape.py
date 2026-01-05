# 鶴亀オセロのサンプル
from excellikeuno import connect_calc
from excellikeuno.drawing import TextShape


(desktop, doc, sheet) = connect_calc()

# Find the first TextShape on the sheet's draw page
text_shape : TextShape = sheet.shapes[0]

# Option 1: use the wrapper property
text_value = text_shape.string

print(text_value)

text_shape.LineStyle = 1  # SOLID
text_shape.LineColor = 0xFF0000  # 赤色に変更
text_shape.LineWidth = 200  # 線の太さを200 (1/100 mm)
text_shape.FillColor = 0xFFFF00  # 黄色に変更
