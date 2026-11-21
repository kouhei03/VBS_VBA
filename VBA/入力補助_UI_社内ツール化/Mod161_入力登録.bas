Attribute VB_Name = "Mod161_入力登録"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 入力専用登録画面提供を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「入力専用登録画面提供」を選択して実行する。

' ── 入力専用登録画面提供 ──
Public Sub 入力専用登録画面提供()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim nameValue As String: nameValue = InputBox("氏名・項目名を入力", "データ登録")
    If Trim$(nameValue) = "" Then Exit Sub
    Dim value As String: value = InputBox("登録値を入力", "データ登録")
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = nameValue
    ws.Cells(r, 2).Value = value
    ws.Cells(r, 3).Value = Now
End Sub




