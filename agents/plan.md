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



