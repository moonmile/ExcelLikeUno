# 将棋盤を作る
from excellikeuno import connect_calc
from excellikeuno.style.border import BorderStyle
from excellikeuno.typing.calc import CellHoriJustify, CellVertJustify, BorderLineStyle

from excellikeuno.typing.structs import BorderLine2
from excellikeuno.typing import Font

(desktop, doc, sheet) = connect_calc()
# sheet.name = "将棋盤"
ban = sheet.range("A1:I9");
ban.CellBackColor = 0xFFFACD  # 背景色を薄い黄色に設定
ban.row_height = 1000  # 行の高さを設定 20 mm
ban.column_width = 1000  # 列の幅を設定 20 mm
# 罫線を設定
for cell in [c for row in ban.cells for c in row]:

    # 罫線の設定 top/left/bottom/right を一括設定
    cell.borders.all.color = 0x000000 # 黒色
    cell.borders.all.weight = 50  # 線の太さを設定
    cell.borders.all.line_style = BorderLineStyle.SOLID  # 0: 実線 
    # BorderStyle の利用
    # cell.borders.all = BorderStyle(color=0x000000, weight=50, line_style=BorderLineStyle.SOLID)


    # センタリング
    cell.HoriJustify = CellHoriJustify.CENTER
    cell.VertJustify = CellVertJustify.CENTER

# フォントの一括変更（内容設定後に適用）
ban.font.size = 16.0
ban.font.color = 0x000000 # 黒色

# Range で一括設定
# ban.borders.all = BorderStyle(color=0x000000, weight=50, line_style=BorderLineStyle.SOLID)

# 駒を配置
pieces = [
    ["香", "桂", "銀", "金", "王", "金", "銀", "桂", "香"],
    ["", "飛", "", "", "", "", "", "角", ""],
    ["歩", "歩", "歩", "歩", "歩", "歩", "歩", "歩", "歩"],
    ["", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", ""],
    ["歩", "歩", "歩", "歩", "歩", "歩", "歩", "歩", "歩"],
    ["", "角", "", "", "", "", "", "飛", ""],
    ["香", "桂", "銀", "金", "王", "金", "銀", "桂", "香"],
]
ban.value = pieces  # 一括で駒を配置

# 相手の駒を反転表示
for r in range(9):
    for c in range(9):
        cell = ban.cell(c, r)
        if pieces[r][c] != "" and r < 3:
            cell.CharRotation = 180  # 180度回転



