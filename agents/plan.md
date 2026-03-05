# クラス構造のリファクタリング

## 目的

- Uno API の .idl から .pyi を生成する
- コーディング時に com.sun.star 以降のクラスを型チェックできるようにする
- XCell, CellProperties をひとまとめにして Cell クラスにする
- cell.borders.all.weight = 50 のように、Excel VBA ライクに構成し直す

## 進め方

### .idl から .pyi を生成

tools/idl_to_pyi.py を実行して、.idl から .pyi を生成する

```bash
python tools/idl_to_pyi.py
```

- offapi/com/ 以下の .idl を利用する
- .pyi にするときは、パッケージ単位にまとめる
- 出力レイアウト方針: 出力ルートを src/stubs とし、IDL のモジュール構造をそのままミラー（例: com/sun/star/sheet/XCell.idl → src/stubs/com/sun/star/sheet/XCell.pyi）。各ディレクトリに __init__.pyi を生成して同階層シンボルを集約する。IDE から解決できるよう PYTHONPATH には src と src/stubs を追加する。
- tools/idl_to_pyi.py は SDK ツール非依存のカスタムパーサで、--idl-root (デフォルト offapi/com) と --out-root (デフォルト src/stubs) を指定可能。生成後、必要なら VS Code をリロードして Pylance のキャッシュをクリアする。
- 足りない .pyi は src/stubs-add/com/... に追加してある XInterface など


### excellikeuno/table/cell.py のリファクタリング

excellikeuno/table/cell.py を sutbs を使ってリファクタリングする
XCell, CellProperties をひとまとめにして Cell クラスにする
- スタブ (src/stubs) を直接 import して型付けする。InterfaceNames はクエリ用定数に限定し、typing/calc からの import は順次削除する。

### Cell/Shape リファクタリング

- cell.py, shape.py をリファクタリングする
- Excel ライクな指定は snake_case に統一する。
- UNO ライクな指定は cell_pros や char_props 経由でアクセスする。
  - UNO クラスを直接利用する場合は props 経由でアクセスする。
  - 複数の UNO クラスを統合している場合は、cell_props や char_props を用意する

Cell の指定

```python
cell = sheet.cell(0, 0)

# Excel ライクな指定
cell.borders.top.color = 0xFF0000
cell.font.size = 12
cell.back_color = 0xFFFF00

# UNO ライクな指定
cell.cell_props.CellBackColor = 0xFFFF00

```

Shape の指定

```python
shape = sheet.shape(0)
# Excel ライクな指定
shape.line_color = 0xFF0000
shape.line_width = 50
# UNO ライクな指定
shape.props.LineColor = 0xFF0000
```

- cell.borders.top.color のプロパティを cell に反映する場合はプロキシ型を使う

プロキシ例

```python
class BorderLineWrapper:
    def __init__(self, bl, on_change):
        self._bl = bl
        self._on_change = on_change

    @property
    def color(self):
        return self._bl.Color

    @color.setter
    def color(self, value):
        self._bl.Color = value
        self._on_change(self._bl)

class BorderSide:
    def _get_wrapper(self):
        bl = self._props.getPropertyValue(self._prop_name)
        return BorderLineWrapper(bl, self._apply)

    def _apply(self, bl):
        self._props.setPropertyValue(self._prop_name, bl)

```

