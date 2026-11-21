Attribute VB_Name = "Mod43_空白行削除"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 空白行を一括削除を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「空白行を一括削除」を選択して実行する。

' ── 空白行を一括削除 ──
Public Sub 空白行を一括削除()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    Dim delRange As Range
    Dim r As Long
    For r = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) = 0 Then
            If delRange Is Nothing Then Set delRange = ws.Rows(r) Else Set delRange = Union(delRange, ws.Rows(r))
        End If
    Next r
    If Not delRange Is Nothing Then delRange.Delete
End Sub




