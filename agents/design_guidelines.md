# 設計方針

UNO API の複雑さを隠し、「Python の 1 クラス = Calc の 1 つの概念」として扱えるようにする。

- UNO オブジェクトを保持するラッパーの基底クラスを作る
- Calc 用のラッパーは、この基底クラスを継承する
- Spreadsheet, Range, Sheet なども同じパターンで統一
- インターフェース名を定数化して IDE 補完を効かせる
- 型ヒントを Protocol で定義して IDE 補完を強化


## コード規約

- UNO API をラップしたときにプロパティ名、メソッド名は UNO ドキュメントに準拠するため PascalCase を使う。
- ただし、Python 側で新規に追加するメソッドやプロパティは snake_case を使う。
- ただし、value, text, formula だけはタイピングしやすいので小文字にしておく。

- UNO API 由来のプロパティは PascalCase を使う。
  - 例: Cell.value, Cell.formula, CellHoriJustify
- Excel VBA ライクなメソッド名を追加する場合は snake_case を使う。
  - 例: sheet.cell(col, row), sheet.range(col1, row1, col2, row2)


## 開発コード例

以下は、開発コード例です。

### UnoObject

```python
class UnoObject:
    def __init__(self, obj):
        self._obj = obj
        self._cache = {}

    def iface(self, name):
        if name not in self._cache:
            self._cache[name] = self._obj.queryInterface(name)
        return self._cache[name]
```

## Cell

```python
class Cell(UnoObject):
    @property
    def value(self):
        return self.iface("com.sun.star.table.XCell").getValue()

    @value.setter
    def value(self, v):
        self.iface("com.sun.star.table.XCell").setValue(v)

    @property
    def formula(self):
        return self.iface("com.sun.star.table.XCell").getFormula()

    @formula.setter
    def formula(self, f):
        self.iface("com.sun.star.table.XCell").setFormula(f)

    @property
    def props(self):
        return self.iface("com.sun.star.beans.XPropertySet")
```

## 利用時の例

```python
import uno
import excellikeuno as elu

(desktop, document, sheet) = uno.connect_to_calc()
cell = sheet.cell(0, 0)  # A1 セルを取得
cell.value = 100
cell.formula = "=A1+B1"

print(cell.CellBackColor) # プロパティで取得

sheet.cell(1, 0).value = 200  # B1 セルを取得して値を設定
sheet.cell(2, 0).formula = "=A1+B1"  # C1 セルに数式を設定

``` 

## CellProperties の公開属性を Cell 直下で提供

CellProperties Service の Public Attributes は Cell のプロパティとして直接アクセスできる（props 経由も可）。主な属性名:

- CellStyle, CellBackColor, IsCellBackgroundTransparent
- HoriJustify, VertJustify, IsTextWrapped, ParaIndent, Orientation, RotateAngle, RotateReference, AsianVerticalMode
- TableBorder, TopBorder, BottomBorder, LeftBorder, RightBorder, NumberFormat, ShadowFormat, CellProtection, UserDefinedAttributes
- DiagonalTLBR, DiagonalBLTR, ShrinkToFit
- TableBorder2, TopBorder2, BottomBorder2, LeftBorder2, RightBorder2, DiagonalTLBR2, DiagonalBLTR2
- CellInteropGrabBag

例: `sheet.cell(0, 0).CellBackColor = 0x112233` で背景色を設定。

# 参照ドキュメント

uno sdk file:///C:/Program%20Files/LibreOffice/sdk/docs/
uno api doc https://api.libreoffice.org/docs/idl/ref/index.html



