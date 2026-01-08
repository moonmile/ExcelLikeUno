# 鶴亀オセロのサンプル
from excellikeuno import connect_calc
from excellikeuno.drawing import TextShape


(desktop, doc, sheet) = connect_calc()

# 1/100 mm 単位で位置・サイズを指定して TextShape を追加
text_shape = sheet.shapes.add_text_shape(1000,1000,13000,2500)
text_shape.string = "Hello, TextShape!"
text_shape.font.size = 24.0
text_shape.font.color = 0x0000FF  # 青色

text_shape.LineStyle = 1  # SOLID
text_shape.LineColor = 0xFF0000  # 赤色に変更
text_shape.LineWidth = 20  # 線の太さを20 (1/100 mm)
text_shape.FillColor = 0xFFFF00  # 黄色に変更
text_shape.HoriJustify = 1  # CENTER
text_shape.VertJustify = 1  # CENTER    
