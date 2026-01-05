# Excel Like UNO

Python ラッパーを通じて LibreOffice Calc の UNO API を操作し、Excel/VBA ライクな操作感を提供します。
Excel マクロからの移行を容易にすることを目的としています。

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

ローカル開発用には本リポジトリをクローンし、`src` を `PYTHONPATH` に通します。

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

ライブラリを以下に配置します。

```powershell
C:\Users\＜ユーザー名＞\AppData\Roaming\LibreOffice\4\user\Scripts\python\
```

以下のように Python スクリプトのパスを通すのと、XSCRIPTCONTEXT を使って接続する connect_calc_script() を用意します。
あと、関数が「マクロ」→「マクロを実行」から見えるように g_exportedScripts に追加しておきます。

```python
import inspect
import os
import sys
from typing import Any, Tuple

# Ensure this script's directory is importable so the local excellike package resolves
BASE_DIR = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)

from excellikeuno.table.sheet import Sheet 

# XSCRIPTCONTEXT に接続する
def connect_calc_script() -> Tuple[Any, Any, Sheet]:
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    controller = doc.getCurrentController()
    sheet = Sheet(controller.getActiveSheet())
    return desktop, doc, sheet

def hello_to_cell():
    ( _, _, sheet ) = connect_calc_script()
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

この手順が実に面倒くさいので、

- BASE_DIR の部分を自動化する（初回だけ設定すればよいらしい）
- connect_calc_script() を excellikeuno.connection.bootstrap モジュールに入れる

この改善を今後検討します。


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

desktop, doc, sheet = connect_calc()

sheet.cell(0, 0).value = 100    # A1
sheet.cell(1, 0).value = 200    # B1
sheet.cell(2, 0).formula = "=A1+B1"  # C1
```

## Writer に接続してテキストを書き込む

```python
from excellikeuno import connect_calc

desktop, doc = connect_writer()
text = doc.text
text.setString("Hello, LibreOffice Writer!")
```

サンプルコードは `samples/` 配下にあり、`xluno.ps1` 経由で実行できます。

```powershell
cd samples
./xluno.ps1 ./calc_sample_mahjong.py
./xluno.ps1 ./writer_sample_text.py
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

0.1.0 (2025-01-05) : 仮リリース

# ライセンス

MIT License

# Author

Tomoaki Masuda (GitHub: @moonmile)

