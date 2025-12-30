# Project Structure

root .
├── README.md
├── LibreOffice-ExcelLike.code-workspace
├── .github/
├── .vscode/
├── agents/                ← AI 用の仕様書・プロンプト
│   ├── tasks.md
│   ├── class_design.md
│   ├── test_execution.md
│   ├── folder_structure.md
│   ├── naming_rules.md
│   ├── operation_spec.md
│   └── design_guidelines.md
├── src/
│   └── excellikeuno/
│       ├── __init__.py
│       ├── connection/
│       │   └── bootstrap.py
│       ├── core/
│       │   ├── base.py
│       │   ├── calc_document.py
│       │   ├── interfaces.py
│       │   └── __init__.py
│       ├── table/
│       │   ├── cell.py
│       │   ├── sheet.py
│       │   └── __init__.py
│       ├── drawing/
│       │   ├── shape.py + shape subclasses (connector, rectangle, ellipse, line, text, polyline/polypolygon, bezier, group, custom, control, measure, page)
│       │   └── __init__.py
│       ├── utils/
│       │   ├── structs.py
│       │   └── __init__.py
│       └── typing/
│           ├── calc.py
│           └── __init__.py
└── tests/
	├── test_cell.py
	├── test_cell_properties.py
	├── test_connect_calc.py
	├── test_shape.py
	├── test_sheet_properties.py
	└── test.ps1