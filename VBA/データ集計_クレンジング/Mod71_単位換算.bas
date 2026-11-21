Attribute VB_Name = "Mod71_単位換算"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 円から千円へ換算を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「円から千円へ換算」を選択して実行する。

' ── 円から千円へ換算 ──
Public Sub 円から千円へ換算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) / 1000
    Next c
End Sub

' ── kgからgへ換算 ──
Public Sub kgからgへ換算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) * 1000
    Next c
End Sub




