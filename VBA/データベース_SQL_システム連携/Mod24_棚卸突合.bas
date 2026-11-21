Attribute VB_Name = "Mod24_棚卸突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 棚卸システム出力実地棚卸表突合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「棚卸システム出力実地棚卸表突合」を選択して実行する。

' ── 棚卸システム出力実地棚卸表突合 ──
Public Sub 棚卸システム出力実地棚卸表突合()
    Dim systemWs As Worksheet: Set systemWs = ActiveSheet
    Dim actualWs As Worksheet
    Dim actualName As String: actualName = InputBox("実地棚卸シート名", "棚卸突合")
    On Error Resume Next
    Set actualWs = ThisWorkbook.Worksheets(actualName)
    On Error GoTo 0
    If actualWs Is Nothing Then Exit Sub
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("商品コード", "システム在庫", "実地在庫")
    Dim r As Long
    For r = 2 To systemWs.Cells(systemWs.Rows.Count, 1).End(xlUp).Row
        output.Cells(r, 1).Value = systemWs.Cells(r, 1).Value
        output.Cells(r, 2).Value = systemWs.Cells(r, 2).Value
        output.Cells(r, 3).Value = actualWs.Cells(r, 2).Value
        If output.Cells(r, 2).Value <> output.Cells(r, 3).Value Then
            output.Rows(r).Interior.Color = RGB(255, 199, 206)
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("棚卸突合")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "棚卸突合"
    End If
End Function




