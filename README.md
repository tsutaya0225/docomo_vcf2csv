# docomo電話帳 VCF to CSV
ドコモ電話帳の編集をしたいがファイルが分散されてて大変……未だ仕様分からず。<br>
ひとまず.vcfと.csvの相互変換を実装。<br>

ShiftJIS → \storage\xxxx-xxxx\SD_PIM\
Unicode → \storage\xxxx-xxxx\com.nttdocomo.android.sdcardbackup\phonebook-utf-8\ <br>
この二つを直しても、ドコモデータコピーアプリがエラーを吐いて元に戻せず。<br>
\storage\xxxx-xxxx\Android\data\com.nttdocomo.android.sdcardbackup\ ← 空っぽ<br>
原因不明<br>

<hr>

面倒なのでPythonに移植中……移植完了……何も改善しない……。
Python版はtkinterdnd2が必要です。
