# クラス設計

## 名前空間

+ core              ; LibreOffice UNO インターフェース定義
    + CalcDocument    ; Calc 文書
    + WriterDocument  ; Writer 文書
+ table             ; Calc 関係
    - Cell
    - Sheet
+ drawing           ; 図形関連
    - Shape
    - ConnectorShape
    - RectangleShape
    - EllipseShape
    - LineShape
    - TextShape
    - ClosedBezierShape
    - OpenBezierShape
    - PolyLineShape
    - PolyPolygonShape
    - PolyPolygonBezierShape
    - GroupShape
    - CustomShape
    - ControlShape
    - MeasureShape
+ chart            ; グラフ関連
    - Chart
+ typing            ; 型ヒント定義
    - BorderLine
    - BorderLine2
    - Position
    - Size
    - Color
    - LineStyle
    - LineDash
    - CellHoriJustify


## クラス構成の比較

### uno での Calc のクラス構造

```
SpreadsheetDocument
 ├─ Sheets
 │    ├─ Sheet
 │    │    ├─ Cells
 │    │    │    ├─ Cell
 │    │    │    ├─ CellRange
 │    │    │    └─ Cursor
 │    │    ├─ Charts
 │    │    ├─ DrawPage
 │    │    ├─ Annotations
 │    │    └─ DataPilotTables
 ├─ NamedRanges
 ├─ DatabaseRanges
 ├─ StyleFamilies
 └─ Controllers
```

### Excel のオブジェクトモデル

```
Application
 ├─ Workbooks
 │    ├─ Workbook
 │    │    ├─ Worksheets
 │    │    │    ├─ Worksheet
 │    │    │    │    ├─ Range
 │    │    │    │    ├─ Cells
 │    │    │    │    ├─ Rows
 │    │    │    │    ├─ Columns
 │    │    │    │    ├─ Shapes
 │    │    │    │    └─ ChartObjects
 │    │    ├─ Names
 │    │    ├─ Charts
 │    │    └─ Windows
 │    └─ ...
 ├─ ActiveWorkbook
 ├─ ActiveSheet
 ├─ Selection（Range / Shape など）
 ├─ Charts
 ├─ AddIns
 └─ Dialogs
```

### Excel Like UNO のクラス構造案

```
CalcDocument
 ├─ Sheets
 │    ├─ Sheet
 │    │    ├─ cell(col, row) : Cell
 │    │    ├─ range(col1, row1, col2, row2) : Range
 │    │    ├─ Rows : TableRows
 │    │    ├─ Columns : TableColumns
 │    │    ├─ Charts : List[Chart]
 │    │    ├─ Shapes : List[Shape]
 │    │    ├─ PivodTables : List[PivodTable]
 │    │    ├─ DrawPage 
 │    │    └─ Annotations
 ├─ NamedRanges
 ├─ StyleFamilies
 ├─ Selection（Range / Shape など）
 ├─ ActiveSheet, ThisSheet
 ActiveDocument, ThisDocument : global

```



## クラス図

```mermaid
classDiagram

    class CalcDocument {
        +Shapes: List[Shape]
        +Charts: List[Chart]
    }

    class Sheet {
        +cell(col: int, row: int) Cell
        +range(col1: int, row1: int, col2: int, row2: int) Range
        +Rows: TableRows
        +Columns: TableColumns
        +Charts: List[Chart]
        +Shapes: List[Shape]
        +font: Font
    }
    class Cell {
        +value: Any
        +formula: str
        +props: XPropertySet
        +cellStyle: string
        +cellBackColor: Color
        +cellForeColor: Color
        +HoriJustify: enum CellHoriJustify
        +font: Font
    }
    class Range : Cell {
        +rows: int
        +columns: int
    }
    class PivodTables {
    }

    class Shape {
        +position: Position
        +size: Size
        +props: XPropertySet
    }
    class ConnectorShape : Shape {
        +startShape: Shape
        +endShape: Shape
        +lineStyle: LineStyle
        +lineDash: LineDash
    }
    class RectangleShape : Shape {
    }
    class EllipseShape : Shape {
    }
    class ClosedBezierShape: Shape {
    }
    class ControlShape : Shape {
    }
    class CustomShape : Shape {
    }
    class EllipseShape : Shape {
    }
    class GroupShape  : Shape {
    }
    class MeasureShape : Shape {
    }
    class OpenBezierShape : Shape {
    }
    clsss PageShape : Shape {
    }
    class PolyLineShape : Shape {
    }
    class PolyPolygonBezierShape : Shape {
    }
    class PolyPolygonShape : Shape {
    }
    class TextShape : Shape {
    }


    struct Font {
        + name : str
        + size : float
        + bold : bool
        + italic : bool
        + underline : bool
        + strikeout : bool
        + color : Color
        + subscript : bool
        + superscript : bool
        + background : Color
        + font_style : int
        + strikthrough : bool
    }
```

