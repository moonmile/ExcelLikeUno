from excellikeuno import connect_calc
from excellikeuno.typing.calc import CellHoriJustify, CellVertJustify
from excellikeuno.typing import Font

(desktop, doc, sheet) = connect_calc() 
cell = sheet.cell(0, 0)  # A1 セルを取得
cell.text = "Hello, World!"  # 値を設定
sheet.range("A1:C1").merge(True)  # A1:C1 を結合

cell.font = Font(name="Arial", size=16, color=0xFF0000)

cell.row_height = 2000  # 行の高さを設定 20 mm
cell.HoriJustify = CellHoriJustify.CENTER
cell.VertJustify = CellVertJustify.CENTER


sheet.cell(0,1).text = "id"
sheet.cell(1,1).text = "name"
sheet.cell(2,1).text = "address"
sheet.range("A2:C2").CellBackColor = 0xFFBF00  # A2:C2 の背景色を設定

data = [
    [1, "masuda", "tokyo"],
    [2, "suzuki", "osaka"],
    [3, "takahashi", "nagoya"],
]
sheet.range("A3:C5").value = data  # 範囲にデータを一括設定


