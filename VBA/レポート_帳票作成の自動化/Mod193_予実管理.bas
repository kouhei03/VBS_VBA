Attribute VB_Name = "Mod193_予実管理"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 予実管理表差異ハイライト自動生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「予実管理表差異ハイライト自動生成」を選択して実行する。

' ── 予実管理表差異ハイライト自動生成 ──
Public Sub 予実管理表差異ハイライト自動生成()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet("予実管理")
    output.Cells.Clear
    source.UsedRange.Copy output.Range("A1")
    Dim lastRow As Long: lastRow = output.Cells(output.Rows.Count, 1).End(xlUp).Row
    output.Cells(1, 4).Value = "差異"
    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(output.Cells(r, 2).Value) And IsNumeric(output.Cells(r, 3).Value) Then
            output.Cells(r, 4).Value = output.Cells(r, 3).Value - output.Cells(r, 2).Value
            If output.Cells(r, 4).Value < 0 Then
                output.Cells(r, 4).Interior.Color = RGB(255, 199, 206)
            Else
                output.Cells(r, 4).Interior.Color = RGB(198, 239, 206)
            End If
        End If
    Next r
    output.Columns.AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = name
    End If
End Function




