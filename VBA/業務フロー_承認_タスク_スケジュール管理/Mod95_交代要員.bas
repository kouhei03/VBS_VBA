Attribute VB_Name = "Mod95_交代要員"
Option Explicit

Public Sub 交代要員調整候補抽出シフト案補助()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:B1").Value = Array("候補者", "資格")
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 3).Value) = "代替可" Then
            ws.Rows(r).Copy output.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("交代要員")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "交代要員"
    End If
End Function
