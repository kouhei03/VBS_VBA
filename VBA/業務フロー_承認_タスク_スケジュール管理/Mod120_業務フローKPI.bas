Attribute VB_Name = "Mod120_業務フローKPI"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 業務フロー完了率期限遵守率差戻率算出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「業務フロー完了率期限遵守率差戻率算出」を選択して実行する。

' ── 業務フロー完了率期限遵守率差戻率算出 ──
Public Sub 業務フロー完了率期限遵守率差戻率算出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim total As Long, completed As Long, deadlineMet As Long, returned As Long
    Dim r As Long
    For r = 2 To lastRow
        total = total + 1
        If CStr(ws.Cells(r, 3).Value) = "完了" Then completed = completed + 1
        If CStr(ws.Cells(r, 4).Value) = "期限遵守" Then deadlineMet = deadlineMet + 1
        If CStr(ws.Cells(r, 5).Value) = "差し戻し" Then returned = returned + 1
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:B1").Value = Array("KPI", "値")
    output.Cells(2, 1).Value = "完了率"
    output.Cells(2, 2).Value = IIf(total > 0, completed / total, 0)
    output.Cells(3, 1).Value = "期限遵守率"
    output.Cells(3, 2).Value = IIf(total > 0, deadlineMet / total, 0)
    output.Cells(4, 1).Value = "差し戻し率"
    output.Cells(4, 2).Value = IIf(total > 0, returned / total, 0)
    output.Range("B2:B4").NumberFormat = "0.0%"
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("業務フローKPI")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "業務フローKPI"
    End If
End Function




