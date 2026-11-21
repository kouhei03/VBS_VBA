Attribute VB_Name = "Mod20_文字列置換"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 特定文字を置換を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「特定文字を置換」を選択して実行する。

' ── 特定文字を置換 ──
Public Function 特定文字を置換(ByVal text As String, ByVal searchText As String, _
                               ByVal replaceText As String) As String
    特定文字を置換 = Replace(text, searchText, replaceText)
End Function




