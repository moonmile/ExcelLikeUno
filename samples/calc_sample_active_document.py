# ActiveDocument, ActiveSheet の利用例

from excellikeuno import connect_calc
from excellikeuno.connection.bootstrap import ActiveCalcDocument
from excellikeuno.core.calc_document import CalcDocument

# connect だけする
connect_calc() 
# アクティブなドキュメントを取得（エイリアス関数を呼び出す）
doc: CalcDocument = ActiveCalcDocument
print(f"Active document title: {doc.title}")
print(f"Active document URL: {doc.url}")
