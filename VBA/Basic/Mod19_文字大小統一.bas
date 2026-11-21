Attribute VB_Name = "Mod19_文字大小統一"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 英字を大文字にを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「英字を大文字に」を選択して実行する。

' ── 英字を大文字に ──
Public Function 英字を大文字に(ByVal text As String) As String
    英字を大文字に = UCase$(text)
End Function

' ── 英字を小文字に ──
Public Function 英字を小文字に(ByVal text As String) As String
    英字を小文字に = LCase$(text)
End Function




