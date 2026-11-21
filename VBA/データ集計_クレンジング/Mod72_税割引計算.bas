Attribute VB_Name = "Mod72_税割引計算"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 消費税込み金額を計算を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「消費税込み金額を計算」を選択して実行する。

' ── 消費税込み金額を計算 ──
Public Sub 消費税込み金額を計算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rate As Double: rate = CDbl(InputBox("税率（例: 10%=0.10）", "消費税", "0.10"))
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) * (1 + rate)
    Next c
End Sub

' ── 割引価格を計算 ──
Public Sub 割引価格を計算()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rate As Double: rate = CDbl(InputBox("割引率（例: 20%=0.20）", "割引計算", "0.20"))
    Dim c As Range
    For Each c In Selection.Cells
        If IsNumeric(c.Value) Then c.Value = CDbl(c.Value) * (1 - rate)
    Next c
End Sub




