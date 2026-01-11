# Project Structure

root .
├── README.md / README.en.md
├── LibreOffice-ExcelLike.code-workspace
├── .github/
├── .vscode/
├── agents/                ← AI 用の仕様書・プロンプト
│   ├── tasks.md
│   ├── class_design.md
│   ├── design_guidelines.md
│   ├── folder_structure.md (this file)
│   ├── naming_rules.md
│   ├── operation_spec.md
│   └── test_execution.md
├── samples/               ← 実行例・デバッグ用スクリプト
│   ├── calc_sample_pie_chart.py ほか calc_*/writer_*/xluno.ps1
│   └── temp_debug.py など検証用
├── python/                ← LibreOffice マクロ配置用コピー
│   └── excellikeuno_macro/ ...（接続・table/drawing 等の同期版）
├── src/
│   └── excellikeuno/
│       ├── connection/    ← 接続・ブートストラップ
│       │   └── bootstrap.py
│       ├── core/          ← UNO ラッパの基盤
│       │   ├── base.py
│       │   ├── calc_document.py
│       │   └── interfaces.py
│       ├── table/         ← Calc シート/セル/レンジ/チャート
│       │   ├── cell.py, cell_properties.py, range.py, sheet.py
│       │   ├── columns.py, rows.py
│       │   └── chart.py   ← Chart/ChartCollection とタイトル/凡例対応
│       ├── drawing/       ← シェイプ群
│       │   ├── shape.py + connector/rectangle/ellipse/line/text/polyline/polypolygon/bezier/group/custom/control/measure/page 等
│       │   └── text_properties, fill/line/shadow properties
│       ├── style/         ← font/character/line/fill などのスタイル
│       ├── utils/         ← 共通ヘルパー（structs など）
│       └── typing/        ← UNO 型定義・構造体名（calc.py, structs.py, interfaces.py）
├── tests/
│   ├── test_cell*.py, test_range*.py, test_shape*.py, test_chart_diagram.py など
│   └── data/              ← テスト用サンプル
└── doc/                   ← 画像などのドキュメント補助