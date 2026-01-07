# 麻雀牌を並べる
from excellikeuno import connect_calc
from excellikeuno.typing.calc import LineStyle
from excellikeuno.typing import Font

(desktop, doc, sheet) = connect_calc()
# sheet.name = "麻雀牌"

hai = ['1','1','1', '2','2','2', '3','3','3', '4','4','4', 'a',' ','a']

# 13個の TextShape をシートに追加
for i in range(hai.__len__()):
    shape = sheet.shapes.add_text_shape(
        x=i * 1600,
        y=1000,
        width=1500,
        height=2000,
        text=str(hai[i]),
        fill_color=0xFFFFE0  # 薄い黄色
    )
    # 背景色を設定
    shape.FillColor = 0xFFFFE0  # 薄い黄色
    # 枠線の色を設定
    shape.LineStyle = LineStyle.SOLID  # SOLID
    shape.LineWidth = 50  # 線の太さを50 (1/100 mm)
    shape.LineColor = 0x008000  # 緑
    # テキストを中央揃え
    shape.HoriJustify = 1  # CENTER
    shape.VertJustify = 1  # CENTER
    # フォント指定（サイズ・フォント名）
    shape.font = Font(size=48.0, name="GL-MahjongTile")

# 麻雀牌を並べる
for i, shape in enumerate(sheet.shapes):
    shape.PositionX = i * 1600  # X位置をずらす
    shape.PositionY = 1000  # Y位置を10 mmに設定

# 完成



