Attribute VB_Name = "Mod21_トレーラ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: レコード件数合計金額トレーラ付与を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「レコード件数合計金額トレーラ付与」を選択して実行する。

' ── レコード件数合計金額トレーラ付与 ──
Public Sub レコード件数合計金額トレーラ付与()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(lastRow + 1, 1).Value = "TRAILER"
    ws.Cells(lastRow + 1, 2).Value = lastRow - 1
    ws.Cells(lastRow + 1, 3).Value = Application.WorksheetFunction.Sum(ws.Columns(3))
End Sub




