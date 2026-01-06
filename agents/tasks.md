# その他のタスク

## Python 実行ファイル
```
C:\Program Files\LibreOffice\program\python'
```

## UNO API ドキュメント
```
C:\Program Files\LibreOffice\sdk\docs\
```

# pip パッケージ関係

token は ~/.pypirc に保存してある。

## インストール
```
& 'C:\Program Files\LibreOffice\program\python' -m pip install excellikeuno
```

## アンインストール
```
& 'C:\Program Files\LibreOffice\program\python' -m pip uninstall excellikeuno
```

## パッケージの作成

```
cd excellikeuno
python -m build
```

## テストサーバーへアップロード
```
twine upload --repository testpypi dist/*
```

## テストサーバーからインストール
```
& 'C:\Program Files\LibreOffice\program\python' -m pip pip install -i https://test.pypi.org/simple excellikeuno
```

## 本番サーバーへアップロード
```
twine upload dist/*
```
