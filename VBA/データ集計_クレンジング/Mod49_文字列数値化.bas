Attribute VB_Name = "Mod49_文字列数値化"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 文字列数値を数値化を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「文字列数値を数値化」を選択して実行する。

' ── 文字列数値を数値化 ──
Public Sub 文字列数値を数値化()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(Trim$(CStr(c.Value))) And Len(Trim$(CStr(c.Value))) > 0 Then
            c.Value = CDbl(Replace(Trim$(CStr(c.Value)), ",", ""))
        End If
    Next c
End Sub




