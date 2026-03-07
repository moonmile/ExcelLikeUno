# クラス構造のリファクタリング

## 目的

- Uno API の .idl から .pyi を生成する
- コーディング時に com.sun.star 以降のクラスを型チェックできるようにする
- XCell, CellProperties をひとまとめにして Cell クラスにする
- cell.borders.all.weight = 50 のように、Excel VBA ライクに構成し直す

## 進め方

### .idl から .pyi を生成

tools/idl_to_pyi.py を実行して、.idl から .pyi を生成する

```bash
python tools/idl_to_pyi.py
```

- offapi/com/ 以下の .idl を利用する
- .pyi にするときは、パッケージ単位にまとめる
- 出力レイアウト方針: 出力ルートを src/stubs とし、IDL のモジュール構造をそのままミラー（例: com/sun/star/sheet/XCell.idl → src/stubs/com/sun/star/sheet/XCell.pyi）。各ディレクトリに __init__.pyi を生成して同階層シンボルを集約する。IDE から解決できるよう PYTHONPATH には src と src/stubs を追加する。
- tools/idl_to_pyi.py は SDK ツール非依存のカスタムパーサで、--idl-root (デフォルト offapi/com) と --out-root (デフォルト src/stubs) を指定可能。生成後、必要なら VS Code をリロードして Pylance のキャッシュをクリアする。
- 足りない .pyi は src/stubs-add/com/... に追加してある XInterface など

