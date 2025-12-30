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



## クラス図

```mermaid
classDiagram

    class CalcDocument {
        +Shapes: List[Shape]
        +Charts: List[Chart]
    }

    class Sheet {
        +cell(col: int, row: int) Cell
        +getCellByPosition(col: int, row: int) Cell
    }
    class Cell {
        +value: Any
        +formula: str
        +props: XPropertySet
        +cellStyle: string
        +cellBackColor: Color
        +cellForeColor: Color
        +HoriJustify: enum CellHoriJustify

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
```

