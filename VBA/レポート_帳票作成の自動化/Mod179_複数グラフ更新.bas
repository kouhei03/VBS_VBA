Attribute VB_Name = "Mod179_複数グラフ更新"
Option Explicit

Public Sub 複数シートグラフ一括更新レイアウト調整()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Dim chart As ChartObject
        For Each chart In ws.ChartObjects
            chart.Chart.Refresh
            chart.Width = 600
            chart.Height = 360
        Next chart
    Next ws
    MsgBox "全シートのグラフを更新しました。", vbInformation
End Sub
