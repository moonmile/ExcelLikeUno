# コード補完チェック用

from excellikeuno import connect_calc
from excellikeuno.style.font import Font
from excellikeuno.style import BorderStyle
from excellikeuno.typing.calc import BorderLineStyle, CellHoriJustify, CellVertJustify, LineStyle

(desktop, doc, sheet) = connect_calc() 

# テキストシェイプ
text_shape = sheet.shapes.add_text_shape(1000,1000,4000,500,text="Sample")
text_shape.font.size = 14
text_shape.font.name = "Courier New"
text_shape.LineStyle = LineStyle.SOLID  # 破線の枠線を設定
text_shape.LineColor = 0x0000FF  # 青色の枠線を設定
text_shape.LineWidth = 50  # 線の太さを設定
text_shape.HoriJustify = 1
text_shape.VertJustify = 1

# 四角
rect_shape = sheet.shapes.add_rectangle_shape(6000,1000,4000,2000)
rect_shape.FillColor = 0xFF00FF  # 紫色の塗りつぶしを設定
rect_shape.LineColor = 0x00FF00  # 緑色の枠線を設定
rect_shape.LineWidth = 30  # 線の太さを設定

rect_shape.line.line_style = LineStyle.DASHED  # 破線の枠線を設定
rect_shape.line.color = 0x00FF00  # 緑色の枠線を設定

