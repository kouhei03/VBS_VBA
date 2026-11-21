Attribute VB_Name = "Mod00_Outlook共通"
Option Explicit

Public Function GetOutlook() As Object
    On Error Resume Next
    Set GetOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0
    If GetOutlook Is Nothing Then Set GetOutlook = CreateObject("Outlook.Application")
End Function

Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function