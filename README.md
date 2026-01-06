# Excel Like UNO

Python ラッパーを通じて LibreOffice Calc の UNO API を操作し、Excel/VBA ライクな操作感を提供します。
Excel マクロからの移行を容易にすることを目的としています。

[Goto English README](README.en.md)

# 主な特徴

- UNO API の複雑さを隠蔽し、Excel/VBA に近いメソッド・プロパティ名で操作可能
- Calc の各概念（シート、セル、範囲、図形など）を Python クラスとしてラップ
- 型定義を充実させ、IDE 補完と静的解析をサポート
- `sheet.cell(col, row).value` のような VBA ライクな書き方でマクロ移行を支援

# 前提環境（Windows）

- LibreOffice 本体（例: `C:\Program Files\LibreOffice`）
- LibreOffice 同梱 Python
	- 実行ファイル: `C:\Program Files\LibreOffice\program\python`
- LibreOffice SDK ドキュメント（任意）
	- UNO API リファレンス: `C:\Program Files\LibreOffice\sdk\docs\`

本リポジトリは Windows 環境で動作確認しています。

# LibreOffice サーバーの起動方法

Calc/Writer へ外部から接続する場合は、先に LibreOffice を「UNO サーバー」として起動します。

```powershell
& "C:\Program Files\LibreOffice\program\soffice" `
	--accept="socket,host=localhost,port=2002;urp;" `
	--norestore --nologo
```

この状態で `connect_calc()` や `connect_writer()` から既存ドキュメントに接続できます。

# インストール

開発中ですが、pip パッケージ化してあります。

```powershell
& 'C:\Program Files\LibreOffice\program\python' -m pip install excellikeuno
```

現在の LibreOffice の Python は 3.11 なので、以下のパスに配置されます。

```powershell
C:\Users\<ユーザー名>\AppData\Roaming\Python\Python311\site-packages\
```

あるいは、ローカル開発用には本リポジトリをクローンし、`src` を `PYTHONPATH` に通します。

```powershell
git clone <this-repo-url>
cd excellikeuno
$env:PYTHONPATH = "${PWD}\src"
```

## LibreOffice を外部から操作する場合

```powershell
$env:PYTHONPATH=＜excellikeunoのパス＞
& 'C:\Program Files\LibreOffice\program\python' ＜スクリプトファイル＞
```

で実行します。

実行する Python が LibreOffice 同梱の Python であることを確認してください。
他の Python 環境では UNO モジュールが見つからず動作しません。

samples/xluno.ps1 のように、自分の環境用にパスを設定しておくと楽です。

```powershell
param(
    [string]$scriptfile = '.'
)
$env:PYTHONPATH='..\src\'
& 'C:\Program Files\LibreOffice\program\python' $scriptfile
```

## LibreOffice 内のマクロで使う場合


pip パッケージでインストールする

```powershell
& 'C:\Program Files\LibreOffice\program\python' -m pip install excellikeuno
```

あるいはライブラリを以下に配置します。

```powershell
C:\Users\＜ユーザー名＞\AppData\Roaming\LibreOffice\4\user\Scripts\python\
```

以下のように Python スクリプトのパスを通すのと、XSCRIPTCONTEXT を使って接続する connect_calc_script() が用意してあります。
あと、関数が「マクロ」→「マクロを実行」から見えるように g_exportedScripts に追加しておきます。

```python
from typing import Any, Tuple
from excellikeuno.table.sheet import Sheet 
from excellikeuno import connect_calc_script

def hello_to_cell():
    ( _, _, sheet ) = connect_calc_script(XSCRIPTCONTEXT)
    sheet.cell(0, 0).text = "Hello Excel Like for Python!"
    sheet.cell(0, 1).text = "こんにちは、Excel Like for Python!"
    sheet.cell(0,0).column_width = 10000  # 幅を設定

    cell = sheet.cell(0,1)
    cell.CellBackColor = 0x006400  # 濃い緑に設定
    cell.CharColor = 0xFFFFFF  # 文字色を白に設定

g_exportedScripts = (
    hello_to_cell,
)
```

![図: マクロの選択](./doc/images/connect_calc_script_macro.jpg)


![図: connect_calc_script の利用](./doc/images/connect_calc_script.jpg)


vscode でコード補完を有効にするために .vscode/settings.json に以下を追加します。

```json
{
    ... 既存の設定

    "python.analysis.autoImportCompletions": true,
    "python.analysis.extraPaths": [
        "C:/Users/masuda/AppData/Roaming/Python/Python311/site-packages"
    ]
}
```


## Linux で使う場合

準備中...

- Linux でのインストールは apt-get などを使えるので比較的楽です。

```bash
sudo apt install libreoffice
sudo apt install python3-uno
```

サーバーの起動

```bash
soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo
```

Linux 版では、ヘッドレス（GUIを使わないモード）がサポートされているので、UNO API 経由での操作に便利です。
これを応用した方法として、WSL や Docker 内で LibreOffice サーバーを動かす方法があります。


## WSL で使う場合

準備中

## Docker コンテナで使う場合

準備中


# 使い方（概要）

## Calc に接続してセルを操作

```python
from excellikeuno import connect_calc
from excellikeuno.typing.calc import CellHoriJustify, CellVertJustify

(desktop, doc, sheet) = connect_calc() 
cell = sheet.cell(0, 0)  # A1 セルを取得
cell.text = "Hello, World!"  # 値を設定
sheet.range("A1:C1").merge(True)  # A1:C1 を結合

cell.font_size = 16
cell.font_name = "Arial"
cell.font_color = 0xFF0000  # フォント色を赤に

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
```

![図1: セル操作](./doc/images/calc_sample_cell.jpg)

## Calc で罫線を引く
```python
from excellikeuno import connect_calc
from excellikeuno.typing.structs import BorderLine
(desktop, doc, sheet) = connect_calc()

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
    cell.font_size = 16.0
    cell.CharColor = 0x000000  # 黒色に設定

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
```

![図2: 将棋盤](./doc/images/calc_sample_shogiban.jpg)


サンプルコードは `samples/` 配下にあり、`xluno.ps1` 経由で実行できます。

```powershell
cd samples
./xluno.ps1 ./calc_sample_cell.py
./xluno.ps1 ./calc_sample_shougiban.py
```

# VS Code での開発

- Python 拡張と PowerShell 拡張を有効化
- テスト実行: コマンドパレットまたは「Run Task」から
	- `Test (LibreOffice Python)` タスクを選択

タスクは LibreOffice 同梱 Python を使って `pytest tests` を実行します。

# テストの仕方

事前に LibreOffice サーバーを起動してから、ルートディレクトリで次を実行します。

```powershell
# サーバー起動
& "C:\Program Files\LibreOffice\program\soffice" --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo

# テスト実行（VS Code タスクと同等）
$env:PYTHONPATH='H:\LibreOffice-ExcelLike\src\'
& 'C:\Program Files\LibreOffice\program\python' -m pytest tests
```

# ドキュメント / UNO API リファレンス

- 本ライブラリの設計・仕様: `agents/` 以下の Markdown
	- クラス設計: `agents/class_design.md`
	- コーディングルール: `agents/coding_rule.md`
	- 設計ガイドライン: `agents/design_guidelines.md`
	- テスト実行手順: `agents/test_execution.md`
- UNO API リファレンス（ローカルインストール）
	- `C:\Program Files\LibreOffice\sdk\docs\`

# バージョン

- 0.1.1 (2025-01-06) : pip パッケージを作成
- 0.1.0 (2025-01-05) : 仮リリース

# ライセンス

MIT License

# Author

Tomoaki Masuda (GitHub: @moonmile)

