# 鶴亀オセロのサンプル
from excellikeuno.connection.bootstrap import connect_calc
from excellikeuno.typing.structs import BorderLine

(desktop, doc, sheet) = connect_calc()
# sheet.name = "鶴亀オセロ"
ban = sheet.range("A1:H8");
ban.CellBackColor = 0x006400  # 背景色を濃い緑色に
ban.row_height = 1000  # 行の高さを設定 10 mm
ban.column_width = 1000  # 列の幅を設定 10 mm
# 罫線を設定
for cell in [c for row in ban.cells for c in row]:
    borderline = BorderLine()
    borderline.Color = 0x000000
    borderline.OuterLineWidth = 50
    borderline.InnerLineWidth = 0
    borderline.LineDistance = 0
    cell.TopBorder = cell.BottomBorder = cell.LeftBorder = cell.RightBorder = borderline

# オセロの駒を Shape で配置
pieces = [
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "黒", "白", "", "", ""],
    ["", "", "", "白", "黒", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", ""],
]

for r in range(8):
    for c in range(8):
        cell = ban.cell(c, r)
        piece = pieces[r][c]
        if piece == "":
            continue
        sheet.shapes.add_ellipse_shape(
            x=cell.position.X + 200,
            y=cell.position.Y + 200,
            width=600,
            height=600,
            fill_color=0x000000 if piece == "黒" else 0xFFFFFF,
            line_color=0x000000,
        )

