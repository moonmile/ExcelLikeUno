# Shape のビットマップを取得する

from excellikeuno import connect_calc
from excellikeuno.style import Fill
from excellikeuno.typing.calc import FillStyle
(desktop, doc, sheet) = connect_calc()

for shape in sheet.shapes:
    if shape.fill.style == FillStyle.BITMAP:
        bitmap_name = shape.fill.bitmap_name
        print(f"Shape '{shape.Name}' has bitmap fill name: {bitmap_name}")

