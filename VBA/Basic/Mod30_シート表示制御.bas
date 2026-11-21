Attribute VB_Name = "Mod30_シート表示制御"
Option Explicit

Public Sub シート表示(ByVal ws As Worksheet)
    ws.Visible = xlSheetVisible
End Sub

Public Sub シート非表示(ByVal ws As Worksheet)
    If wbVisibleSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 3001, , "表示中の最後のシートは非表示にできません。"
    End If
    ws.Visible = xlSheetHidden
End Sub

Public Sub シートを非常に隠す(ByVal ws As Worksheet)
    If wbVisibleSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 3002, , "表示中の最後のシートをVeryHiddenにはできません。"
    End If
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function wbVisibleSheetCount(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then wbVisibleSheetCount = wbVisibleSheetCount + 1
    Next ws
End Function
