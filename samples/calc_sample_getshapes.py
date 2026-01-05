# アクティブシートの図形を取得するサンプル
from excellikeuno import connect_calc
(desktop, doc, sheet) = connect_calc()
# sheet.name = "図形取得サンプル"
# シート上の全図形を列挙
for shape in sheet.shapes:
    print(f"Shape: {shape}, Type: {type(shape)}")
    if hasattr(shape, "string"):
        print(f"  Text: {shape.string}")
    print(f"  Position: ({shape.Position.X}, {shape.Position.Y}), Size: ({shape.Size.Width} x {shape.Size.Height})")
    print(f"  FillColor: {hex(shape.FillColor)}, LineColor: {hex(shape.LineColor)}")
    print(f"  LineStyle: {shape.LineStyle}, LineWidth: {shape.LineWidth}")

# 図形の位置やサイズ、色などのプロパティにアクセス可能


