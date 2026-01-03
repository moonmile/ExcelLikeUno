# 鶴亀オセロのサンプル
from excellikeuno.connection.bootstrap import connect_calc
from excellikeuno.drawing import TextShape


(desktop, doc, sheet) = connect_calc()

# Find the first TextShape on the sheet's draw page
text_shape : TextShape = sheet.shapes[0]

# Option 1: use the wrapper property
text_value = text_shape.string

print(text_value)
