# Excel Like UNO

Provides an Excel/VBA-like programming experience to LibreOffice Calc through a Python wrapper over the UNO API. The goal is to make migration from Excel macros easier.

[Goto 日本語 README](README.md)

# Key Features

- Hide the complexity of the UNO API and expose methods/properties close to Excel/VBA
- Wrap Calc concepts (sheets, cells, ranges, shapes, etc.) as Python classes
- Ship rich type hints to support IDE completion and static analysis
- Support VBA-like code such as `sheet.cell(col, row).value` to help migrate macros

# Design Principles

- One wrapper class per Calc concept, all built on a shared `UnoObject` cache of queried interfaces
- UNO-derived names stay PascalCase; newly added Python helpers use snake_case (shortcuts like `value`, `formula`, `text` kept lowercase)
- Prefer interface-name constants and Protocol-based type hints to improve IDE completion and avoid typos
- Provide Excel/VBA-like convenience such as `sheet.cell(col, row)` and `range.borders` to simplify migration

# Project Layout

- `src/excellikeuno/`: library code (connection/bootstrap, core base, Calc/table wrappers, drawing/shapes, utils, typing)
- `samples/`: runnable examples; use `samples/xluno.ps1` with the LibreOffice-bundled Python
- `tests/`: pytest cases targeting the LibreOffice runtime (UNO server required)
- `agents/`: design docs and prompts (single source of truth for architecture and naming)

# Prerequisites (Windows)

- LibreOffice installation (e.g. `C:\Program Files\LibreOffice`)
- LibreOffice-bundled Python
  - Executable: `C:\Program Files\LibreOffice\program\python`
- LibreOffice SDK documentation (optional)
  - UNO API reference: `C:\Program Files\LibreOffice\sdk\docs\`

This repository is currently tested on Windows.

# Starting the LibreOffice Server

When controlling Calc/Writer from an external process, first start LibreOffice as a "UNO server":

```powershell
& "C:\Program Files\LibreOffice\program\soffice" `
  --accept="socket,host=localhost,port=2002;urp;" `
  --norestore --nologo
```

With this server running, you can connect from `connect_calc()` or `connect_writer()` to an existing document.

# Installation

A pip package is available (under development):

```powershell
& 'C:\Program Files\LibreOffice\program\python' -m pip install excellikeuno
```

LibreOffice currently bundles Python 3.11, so the package is installed under a user site-packages directory similar to:

```powershell
C:\Users\<UserName>\AppData\Roaming\Python\Python311\site-packages\
```

For local development, clone this repository and add `src` to `PYTHONPATH` instead of using the installed package:

```powershell
git clone <this-repo-url>
cd excellikeuno
$env:PYTHONPATH = "${PWD}\src"
```

## Using LibreOffice from an external script

```powershell
$env:PYTHONPATH=<path-to-excellikeuno>
& 'C:\Program Files\LibreOffice\program\python' <script-file>
```

Make sure that the Python executable is the one bundled with LibreOffice. Other Python environments usually do not have the UNO modules installed and will fail.

It is convenient to have a helper script like `samples/xluno.ps1` adjusted to your environment:

```powershell
param(
    [string]$scriptfile = '.'
)
$env:PYTHONPATH='..\src\'
& 'C:\Program Files\LibreOffice\program\python' $scriptfile
```

## Using from LibreOffice internal macros

First, install the package with the LibreOffice-bundled Python:

```powershell
& 'C:\Program Files\LibreOffice\program\python' -m pip install excellikeuno
```

Alternatively, place this library directly under:

```powershell
C:\Users\<UserName>\AppData\Roaming\LibreOffice\4\user\Scripts\python\
```

The library provides `connect_calc_script()`, which uses `XSCRIPTCONTEXT` to connect to the active Calc document. Add your macro function to `g_exportedScripts` so it is visible under "Tools" → "Macros" → "Run Macro":

```python
from typing import Any, Tuple
from excellikeuno.table.sheet import Sheet 
from excellikeuno import connect_calc_script

def hello_to_cell():
    (_, _, sheet) = connect_calc_script(XSCRIPTCONTEXT)
    sheet.cell(0, 0).text = "Hello Excel Like for Python!"
    sheet.cell(0, 1).text = "こんにちは、Excel Like for Python!"
    sheet.cell(0,0).column_width = 10000  # set width

    cell = sheet.cell(0,1)
    cell.CellBackColor = 0x006400  # dark green
    cell.CharColor = 0xFFFFFF  # white text

g_exportedScripts = (
    hello_to_cell,
)
```

![Macro selection](./doc/images/connect_calc_script_macro.jpg)

![Using connect_calc_script](./doc/images/connect_calc_script.jpg)

To enable code completion in VS Code, add the following to `.vscode/settings.json` (adjust the path to your user name and Python version):

```json
{
    // existing settings...

    "python.analysis.autoImportCompletions": true,
    "python.analysis.extraPaths": [
        "C:/Users/masuda/AppData/Roaming/Python/Python311/site-packages"
    ]
}
```

## Using on Linux

Work in progress, but installation on Linux is relatively straightforward using distribution packages.

```bash
sudo apt install libreoffice
sudo apt install python3-uno
```

Start the UNO server:

```bash
soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo
```

On Linux, headless (no GUI) mode is available and convenient when controlling LibreOffice via the UNO API. You can also run the LibreOffice server inside WSL or Docker.

For Python macros, install the package into the OS Python environment. Example on Ubuntu 20.04 + Python 3.12 (create any working folder such as `~/libre`):

```bash
mkdir libre
cd libre
python -m venv .venv
.venv/bin/pip install excellikeuno
ls .venv/lib/python3.12/site-packages
sudo cp -r ~/libre/.venv/lib/python3.12/site-packages/excellikeuno /usr/lib/python3/dist-packages/
```

## Using with WSL

Work in progress

## Using with Docker containers

Work in progress

# Usage Overview

## Connect to Calc and work with cells

```python
from excellikeuno import connect_calc
from excellikeuno.typing.calc import CellHoriJustify, CellVertJustify

(desktop, doc, sheet) = connect_calc() 
cell = sheet.cell(0, 0)  # A1
cell.text = "Hello, World!"
sheet.range("A1:C1").merge(True)

cell.font.size = 16
cell.font.name = "Arial"
cell.font.color = 0xFF0000  # red text

cell.row_height = 2000  # 20 mm
cell.HoriJustify = CellHoriJustify.CENTER
cell.VertJustify = CellVertJustify.CENTER

sheet.cell(0,1).text = "id"
sheet.cell(1,1).text = "name"
sheet.cell(2,1).text = "address"
sheet.range("A2:C2").CellBackColor = 0xFFBF00  # header background

data = [
    [1, "masuda", "tokyo"],
    [2, "suzuki", "osaka"],
    [3, "takahashi", "nagoya"],
]
sheet.range("A3:C5").value = data  # bulk assign
```

![Cell operations](./doc/images/calc_sample_cell.jpg)

## Draw borders in Calc

```python
from excellikeuno import connect_calc
from excellikeuno.typing.calc import CellHoriJustify, CellVertJustify, BorderLineStyle

(desktop, doc, sheet) = connect_calc()

ban = sheet.range("A1:I9")
ban.CellBackColor = 0xFFFACD  # light yellow
ban.row_height = 1000
ban.column_width = 1000

for cell in [c for row in ban.cells for c in row]:
    cell.borders.all.color = 0x000000
    cell.borders.all.weight = 50
    cell.borders.all.line_style = BorderLineStyle.SOLID
    cell.HoriJustify = CellHoriJustify.CENTER
    cell.VertJustify = CellVertJustify.CENTER

ban.font.size = 16.0
ban.font.color = 0x000000

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
ban.value = pieces

for r in range(9):
    for c in range(9):
        cell = ban.cell(c, r)
        if pieces[r][c] != "" and r < 3:
            cell.CharRotation = 180
```

![Shogi board](./doc/images/calc_sample_shogiban.jpg)

Sample code is under `samples/` and can be run via `xluno.ps1`:

```powershell
cd samples
./xluno.ps1 ./calc_sample_cell.py
./xluno.ps1 ./calc_sample_shogiban.py
./xluno.ps1 ./writer_sample_text.py
```

# Development with VS Code

- Enable the Python extension and PowerShell extension
- To run tests, use the Command Palette or "Run Task" and choose `Test (LibreOffice Python)`

The task uses the LibreOffice-bundled Python to run `pytest tests`.

Recommended flow:
- Start the UNO server (see below)
- Ensure `PYTHONPATH` includes `src`
- Run the `Test (LibreOffice Python)` task or the equivalent command below

# Running Tests

Start the LibreOffice server first, then from the repository root run:

```powershell
# Start server
& "C:\Program Files\LibreOffice\program\soffice" --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo

# Run tests (same as the VS Code task)
$env:PYTHONPATH='H:\LibreOffice-ExcelLike\src\'
& 'C:\Program Files\LibreOffice\program\python' -m pytest tests
```

# Documentation / UNO API Reference

- Project design and specs live under `agents/` (update these first when changing behavior):
  - `agents/class_design.md` (class responsibilities)
  - `agents/design_guidelines.md` (wrapping patterns and naming rules)
  - `agents/folder_structure.md` (module layout)
  - `agents/naming_rules.md` (Calc naming conventions; stub)
  - `agents/operation_spec.md` (Calc operation spec; stub)
  - `agents/tasks.md` (work items for agents)
  - `agents/test_execution.md` (how to run tests/headless modes)
- UNO API reference (local installation):
  - `C:\Program Files\LibreOffice\sdk\docs\`

## Blog

- [Using ExcelLikeUno with LibreOffice Calc on Linux | Moonmile Solutions Blog](https://www.moonmile.net/blog/archives/11933)

# Version

- 0.2.0 (2025-01-09): Added `font` on Cell/Range/Shape and `borders` on Cell/Range
- 0.1.1 (2025-01-06): Built pip package
- 0.1.0 (2025-01-05): Pre-release

# License

MIT License

# Author

Tomoaki Masuda (GitHub: @moonmile)
