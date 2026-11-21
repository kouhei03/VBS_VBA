Attribute VB_Name = "Mod219_BeforeAfter突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 改修前後集計計算結果突合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「改修前後集計計算結果突合」を選択して実行する。

' ── 改修前後集計計算結果突合 ──
Public Sub 改修前後集計計算結果突合()
    Dim beforeWs As Worksheet: Set beforeWs = ActiveSheet
    Dim afterName As String: afterName = InputBox("改修後シート名", "Before/After")
    Dim afterWs As Worksheet
    On Error Resume Next
    Set afterWs = ThisWorkbook.Worksheets(afterName)
    On Error GoTo 0
    If afterWs Is Nothing Then Exit Sub
    Dim report As Worksheet: Set report = GetOrCreateSheet
    report.Cells.Clear
    report.Range("A1:D1").Value = Array("セル", "Before", "After", "差分")
    Dim r As Long
    For r = 1 To beforeWs.UsedRange.Rows.Count
        report.Cells(r + 1, 1).Value = "R" & r
        report.Cells(r + 1, 2).Value = beforeWs.Cells(r, 1).Value
        report.Cells(r + 1, 3).Value = afterWs.Cells(r, 1).Value
        If IsNumeric(beforeWs.Cells(r, 1).Value) And IsNumeric(afterWs.Cells(r, 1).Value) Then
            report.Cells(r + 1, 4).Value = CDbl(afterWs.Cells(r, 1).Value) - CDbl(beforeWs.Cells(r, 1).Value)
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("BeforeAfter差分")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "BeforeAfter差分"
    End If
End Function




