Attribute VB_Name = "Mod28_シート存在確認"
Option Explicit

Public Function シート存在確認(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    シート存在確認 = Not ws Is Nothing
End Function
