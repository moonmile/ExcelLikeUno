"""LibreOffice Calc sample macro."""
from re import X
from typing import Any, Tuple
from excellikeuno import connect_calc_script
from excellikeuno.sheet.spreadsheet import Spreadsheet 

def hello_to_cell():
    ( _, _, sheet ) = connect_calc_script(XSCRIPTCONTEXT)
    sheet.cell(0, 0).text = "Hello Excel Like for Python!"
    sheet.cell(0, 1).text = "こんにちは、Excel Like for Python!"
    sheet.cell(0,0).column_width = 10000  # 幅を設定

    cell = sheet.cell(0,1)
    cell.backcolor = 0x006400  # 濃い緑に設定
    cell.color = 0xFFFFFF  # 文字色を白に設定

g_exportedScripts = (
    hello_to_cell,
)
