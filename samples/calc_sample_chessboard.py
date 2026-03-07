# sample chessboard
import uno
from excellikeuno import connect_calc
from excellikeuno.drawing.shape import Shape 
from excellikeuno.sheet import Cell

(desktop, doc, sheet) = connect_calc()

# sheet.name = "chess board"

ban = sheet.range("A1:H8")
ban.row_height = 1000  # row height 10 mm
ban.column_width = 1000  # column width 10 mm
colors = [0xFFFFFF, 0x000000]
for r in range(8):
    for c in range(8):
        cell = ban.cell(c, r)
        # cell.CellBackColor = colors[(r + c) % 2]
        cell.horizontal_align = 2 # CENTER
        cell.vertical_align = 2 # CENTER
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

# make RectangleShape, display to background by bitmap
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

    x = cell.position.X
    y = cell.position.Y
    w = cell.column_width
    h = cell.row_height
    shape_copy = sheet.shapes.add_rectangle_shape(x,y, w, h )
    shape_copy.fill.style = shape.fill.style
    shape_copy.fill.bitmap_name = shape.fill.bitmap_name
    # set background by dispatcher call
    shape_copy.to_background()

    return shape_copy

# set shape_white and shape_back to sheet.range("A1:H8") 
for r in range(8):
    for c in range(8):
        cell = ban.cell(c, r)
        if (r + c) % 2 == 0:
            copy_behind(cell, shape_white)
        else:
            copy_behind(cell, shape_black)

# delete the original shape_white and shape_black
sheet.shapes.remove(shape_white)
sheet.shapes.remove(shape_black)


