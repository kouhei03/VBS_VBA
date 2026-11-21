Attribute VB_Name = "Mod141_グラフPPT貼付"
Option Explicit

Public Sub ExcelグラフをPowerPointへ一括貼付()
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    Dim presentation As Object: Set presentation = ppt.Presentations.Add
    Dim chart As ChartObject
    For Each chart In ActiveSheet.ChartObjects
        chart.Copy
        Dim slide As Object: Set slide = presentation.Slides.Add(presentation.Slides.Count + 1, 1)
        slide.Shapes.Paste
    Next chart
    MsgBox "グラフを貼り付けました。", vbInformation
End Sub
