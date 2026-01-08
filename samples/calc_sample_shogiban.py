# 将棋盤を作る
from excellikeuno import connect_calc
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
    borderline = BorderLine2()
    borderline.Color = 0x000000
    borderline.OuterLineWidth = 50
    borderline.InnerLineWidth = 0
    borderline.LineDistance = 0
    # borderline.LineStyle =  BorderLineStyle.SOLID  # solid

    # 一括設定（新しい Border プロキシ経由）
    cell.border.all = borderline

    # センタリング
    cell.HoriJustify = CellHoriJustify.CENTER
    cell.VertJustify = CellVertJustify.CENTER

# フォントの一括変更（内容設定後に適用）
ban.font.size = 16.0
ban.font.color = 0x000000 # 青色

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



