# Writer に書き込むテスト

from excellikeuno.connection.bootstrap import connect_writer
(desktop, doc) = connect_writer()

# ドキュメントの最初のテキストコンテナを取得
text_container = doc.text
# テキストを追加
text_container.setString("Hello, LibreOffice Writer!")
# 改行を追加
text_container.insertControlCharacter(
    text_container.getEnd(), 0, False
)
# さらにテキストを追加
text_container.insertString(
    text_container.getEnd(), "This is a sample text added via UNO.", False
)


