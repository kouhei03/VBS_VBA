Attribute VB_Name = "Mod139_PowerPoint起動"
Option Explicit

Public Sub PowerPoint起動新規作成()
    Dim ppt As Object
    On Error Resume Next
    Set ppt = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If ppt Is Nothing Then Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    ppt.Presentations.Add
End Sub
