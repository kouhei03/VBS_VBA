Attribute VB_Name = "Mod140_PowerPointスライド"
Option Explicit

Public Sub ExcelデータからPowerPointスライド自動生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    Dim presentation As Object: Set presentation = ppt.Presentations.Add
    Dim r As Long
    For r = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim slide As Object: Set slide = presentation.Slides.Add(presentation.Slides.Count + 1, 1)
        slide.Shapes(1).TextFrame.TextRange.Text = CStr(ws.Cells(r, 1).Value)
        If ws.Cells(r, 2).Value <> "" Then
            slide.Shapes(2).TextFrame.TextRange.Text = CStr(ws.Cells(r, 2).Value)
        End If
    Next r
    MsgBox "スライドを生成しました。", vbInformation
End Sub
