# チェスボードのサンプル
import uno
from excellikeuno import connect_calc
from excellikeuno.drawing.shape import Shape 
from excellikeuno.table.cell import Cell
from com.sun.star.beans import PropertyValue  # type: ignore
(desktop, doc, sheet) = connect_calc()
# sheet.name = "チェスボード"
ban = sheet.range("A1:H8")
ban.row_height = 1000  # 行の高さを設定 10 mm
ban.column_width = 1000  # 列の幅を設定 10 mm
colors = [0xFFFFFF, 0x000000]
for r in range(8):
    for c in range(8):
        cell = ban.cell(c, r)
        # cell.CellBackColor = colors[(r + c) % 2]
        cell.HoriJustify = 2 # CENTER
        cell.VertJustify = 2 # CENTER
        piece = ""
        if r == 0 or r == 7:
            piece = "♜♞♝♛♚♝♞♜"[c]
        elif r == 1 or r == 6:
            piece = "♟"[0] 
        cell.value = piece
        cell.font.size = 20
        cell.font.name = "Arial Unicode MS"
        if 0 <= r <= 2:
            cell.font.color = 0xFFFFFF

# RectangleShape を作成し、背景にビットマップを表示
shape_white = sheet.shapes.add_rectangle_shape(
    x=0,
    y=0,
    width=1000,
    height=1000 )
shape_white.fill.bitmap_name = "Concrete"

shape_black = sheet.shapes.add_rectangle_shape(
    x=1000,
    y=1000,
    width=1000,
    height=1000 )
shape_black.fill.bitmap_name = "Parchment Paper"

def copy_behind(cell : Cell, shape: Shape):

    # cell の位置を取得
    # shape をコピー
    # cell の位置に shape を移動する
    # 背景に移動する
    x = cell.position.X
    y = cell.position.Y
    w = cell.column_width
    h = cell.row_height
    shape_copy = sheet.shapes.add_rectangle_shape(x,y, w, h )
    shape_copy.fill.style = shape.fill.style
    shape_copy.fill.bitmap_name = shape.fill.bitmap_name
    # UI の「配置→背景へ」と同等の dispatcher 呼び出し
    shape_copy.to_background()

    return shape_copy

# shape_white と shape_back を sheet.range("A1:H8") にコピーしてい配置
for r in range(8):
    for c in range(8):
        cell = ban.cell(c, r)
        if (r + c) % 2 == 0:
            copy_behind(cell, shape_white)
        else:
            copy_behind(cell, shape_black)

# 元の shape_white と shape_black を削除
sheet.shapes.remove(shape_white)
sheet.shapes.remove(shape_black)


