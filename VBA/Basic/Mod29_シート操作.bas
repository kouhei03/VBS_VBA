Attribute VB_Name = "Mod29_シート操作"
Option Explicit

Public Function シート追加(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Set シート追加 = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    シート追加.Name = sheetName
End Function

Public Sub シート名称変更(ByVal ws As Worksheet, ByVal newName As String)
    ws.Name = newName
End Sub

Public Sub シート削除(ByVal ws As Worksheet)
    If wbSheetCount(ws.Parent) <= 1 Then
        Err.Raise vbObjectError + 2901, , "ブック内の最後のシートは削除できません。"
    End If
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

Private Function wbSheetCount(ByVal wb As Workbook) As Long
    wbSheetCount = wb.Worksheets.Count
End Function
