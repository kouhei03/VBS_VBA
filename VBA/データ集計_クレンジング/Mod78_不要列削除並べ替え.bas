Attribute VB_Name = "Mod78_不要列削除並べ替え"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 不要列を削除を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「不要列を削除」を選択して実行する。

' ── 不要列を削除 ──
Public Sub 不要列を削除()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim colText As String: colText = InputBox("削除する列番号（カンマ区切り）", "列削除")
    If colText = "" Then Exit Sub
    Dim parts As Variant: parts = Split(colText, ",")
    Dim i As Long
    For i = UBound(parts) To LBound(parts) Step -1
        ws.Columns(CLng(Trim$(CStr(parts(i))))).Delete
    Next i
End Sub




