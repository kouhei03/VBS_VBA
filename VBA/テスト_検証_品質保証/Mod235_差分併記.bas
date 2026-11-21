Attribute VB_Name = "Mod235_差分併記"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 差分レポート前後値差分併記を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「差分レポート前後値差分併記」を選択して実行する。

' ── 差分レポート前後値差分併記 ──
Public Sub 差分レポート前後値差分併記()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet
    out.Cells.Clear
    out.Range("A1:D1").Value = Array("セル", "改修前", "改修後", "差分")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        out.Cells(r, 1).Value = ws.Cells(r, 1).Address
        out.Cells(r, 2).Value = ws.Cells(r, 1).Value
        out.Cells(r, 3).Value = ws.Cells(r, 2).Value
        If IsNumeric(ws.Cells(r, 1).Value) And IsNumeric(ws.Cells(r, 2).Value) Then
            out.Cells(r, 4).Value = CDbl(ws.Cells(r, 2).Value) - CDbl(ws.Cells(r, 1).Value)
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("差分レポート")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "差分レポート"
    End If
End Function




