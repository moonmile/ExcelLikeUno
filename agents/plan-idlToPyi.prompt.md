## Plan: IDL→PYI 生成（継承対応付き）

offapi 全体の IDL から .pyi を生成し、src/sutubs/com/... にパッケージ構造を保ったスタブを出力する。interface/struct/enum・定数に加え、IDL で宣言された継承関係（親インターフェース／サービス拡張）を Protocol/MRO に反映して IDE 補完を強化する。

**Steps**
1. 既存型定義の流儀確認: [src/excellikeuno/typing](src/excellikeuno/typing) の定数命名・Protocol 形・継承有無を把握。
2. 出力レイアウト決定: IDL のパッケージ構造を src/sutubs/com/... に .pyi で再現し、各パッケージに __init__.pyi で同階層シンボルを集約する方針を [agents/plan.md](agents/plan.md) に追記。
3. パーサ実装: [tools/idl_to_pyi.py](tools/idl_to_pyi.py) を作成し [offapi/com](offapi/com) を走査。interface/struct/enum・定数に加え、IDL の `: super` 句や `service extends` 句を解析し、生成 Protocol の継承リストに親を反映する。
4. マッピング規約: interface→Protocol（属性/メソッド＋親列挙）、struct→TypedDict/Protocol 形（必要なら基底 struct 継承を反映）、enum/定数→モジュール定数。UNO 側の PascalCase を基本保持し、既知の小文字例外のみ適用。
5. 依存・継承解決: IDL の include/import を解析し、親インターフェース／拡張サービスを適切に import しつつ相対/絶対 import を生成。
6. CLI 追加: tools/idl_to_pyi.py に source/root/overwrite などのオプションを付け、デフォルト出力を src/sutubs に設定し、決定論的順序で生成。
7. 試験生成: スクリプト実行でスタブ生成後、Calc 主要インターフェース（例: XCell が XCellRange や XPropertySet を継承しているか）を目視確認し、継承と属性が正しく出ているか検証。
8. 型検証: PYTHONPATH を src と src/sutubs に向け、mypy/pyright をサンプル/ラッパーに当てて解決確認。親子関係が補完に反映されることを確認。
9. ドキュメント更新: 実行方法・オプション・出力レイアウトと継承反映の仕様を [agents/plan.md](agents/plan.md) に追記し、SDK ツール非依存であることを明記。

**Verification**
- `python tools/idl_to_pyi.py` 実行で src/sutubs/com/... がエラーなく生成され、Protocol に親インターフェースが含まれること。
- 生成スタブの XCell/CellProperties など主要型で継承と属性/メソッドが期待どおり出力されていることを目視確認。
- mypy/pyright を既存ラッパー/サンプルに走らせ、継承関係を含め型解決が通ることを確認。

**Decisions**
- スコープ: offapi 全体を対象
- 出力先: src/sutubs/com/... に IDL 構造をミラー
- 生成手段: LibreOffice SDK ツール不使用のカスタムパーサ
- カバー範囲: interface, struct, enum/定数（継承を反映）
