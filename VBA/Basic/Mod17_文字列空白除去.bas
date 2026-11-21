Attribute VB_Name = "Mod17_文字列空白除去"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 前後空白を除去を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「前後空白を除去」を選択して実行する。

' ── 前後空白を除去 ──
Public Function 前後空白を除去(ByVal text As Variant) As String
    If IsError(text) Or IsNull(text) Or IsEmpty(text) Then
        前後空白を除去 = ""
    Else
        前後空白を除去 = Trim$(CStr(text))
    End If
End Function




