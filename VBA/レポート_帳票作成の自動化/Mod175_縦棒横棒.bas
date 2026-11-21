Attribute VB_Name = "Mod175_縦棒横棒"
Option Explicit

Public Sub 縦棒横棒グラフ自動作成()
    Dim dataRange As Range
    Set dataRange = Application.InputBox("グラフ化するセル範囲を選択", "グラフ作成", Selection.Address, , , , , 8)
    Dim mode As String: mode = InputBox("種類: 1=縦棒、2=横棒", "グラフ作成", "1")
    Dim chartType As XlChartType
    If mode = "1" Then chartType = xlColumnClustered Else chartType = xlBarClustered
    ActiveSheet.Shapes.AddChart2(201, chartType).Chart.SetSourceData dataRange
End Sub
