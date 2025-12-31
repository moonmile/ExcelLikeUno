from excellikeuno.connection.bootstrap import connect_calc

(desktop, doc, sheet) = connect_calc() 
cell = sheet.cell(0, 0)  # A1 セルを取得
cell.text = "Hello, World!"  # 値を設定
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






