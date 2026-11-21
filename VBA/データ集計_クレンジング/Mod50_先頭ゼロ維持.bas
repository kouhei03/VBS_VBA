Attribute VB_Name = "Mod50_先頭ゼロ維持"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 数値コードを文字列化してゼロ維持を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「数値コードを文字列化してゼロ維持」を選択して実行する。

' ── 数値コードを文字列化してゼロ維持 ──
Public Sub 数値コードを文字列化してゼロ維持()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then
            c.NumberFormat = "@"
            c.Value = Format$(CDec(c.Value), "0")
        End If
    Next c
End Sub




