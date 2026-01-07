# 将棋盤を作る
from excellikeuno import connect_calc
from excellikeuno.typing.calc import CellHoriJustify, CellVertJustify
from excellikeuno.typing.structs import BorderLine
from excellikeuno.typing import Font

(desktop, doc, sheet) = connect_calc()
# sheet.name = "将棋盤"
ban = sheet.range("A1:I9");
ban.CellBackColor = 0xFFFACD  # 背景色を薄い黄色に設定
ban.row_height = 1000  # 行の高さを設定 20 mm
ban.column_width = 1000  # 列の幅を設定 20 mm
# 罫線を設定
for cell in [c for row in ban.cells for c in row]:
    borderline = BorderLine()
    borderline.Color = 0x000000
    borderline.OuterLineWidth = 50
    borderline.InnerLineWidth = 0
    borderline.LineDistance = 0

    cell.TopBorder = borderline
    cell.BottomBorder = borderline
    cell.LeftBorder = borderline
    cell.RightBorder = borderline
    # センタリング
    cell.HoriJustify = CellHoriJustify.CENTER
    cell.VertJustify = CellVertJustify.CENTER
    # フォントサイズを大きく
    # cell.font = Font(size=16.0, color=0x000000)
    cell.font.size = 16.0
    cell.font.color = 0x000000

    

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




