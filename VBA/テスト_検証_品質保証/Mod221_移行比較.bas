Attribute VB_Name = "Mod221_移行比較"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: バージョンアップ移行前後件数合計件数内訳比較を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「バージョンアップ移行前後件数合計件数内訳比較」を選択して実行する。

' ── バージョンアップ移行前後件数合計件数内訳比較 ──
Public Sub バージョンアップ移行前後件数合計件数内訳比較()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:D1").Value = Array("項目", "移行前", "移行後", "差分")
    report.Cells(2, 1).Value = "件数"
    report.Cells(2, 2).Value = ws.Cells(2, 1).Value
    report.Cells(2, 3).Value = ws.Cells(2, 2).Value
    report.Cells(2, 4).Value = CDbl(ws.Cells(2, 2).Value) - CDbl(ws.Cells(2, 1).Value)
    report.Cells(3, 1).Value = "合計金額"
    report.Cells(3, 2).Value = ws.Cells(3, 1).Value
    report.Cells(3, 3).Value = ws.Cells(3, 2).Value
    report.Cells(3, 4).Value = CDbl(ws.Cells(3, 2).Value) - CDbl(ws.Cells(3, 1).Value)
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("移行比較")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "移行比較"
    End If
End Function




