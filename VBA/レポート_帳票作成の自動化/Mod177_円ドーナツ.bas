Attribute VB_Name = "Mod177_円ドーナツ"
Option Explicit

Public Sub 円ドーナツグラフ自動作成()
    Dim dataRange As Range
    Set dataRange = Application.InputBox("グラフ化する範囲を選択", "円グラフ", Selection.Address, , , , , 8)
    Dim mode As String: mode = InputBox("種類: 1=円、2=ドーナツ", "円グラフ", "1")
    Dim chartType As XlChartType
    If mode = "1" Then chartType = xlPie Else chartType = xlDoughnut
    ActiveSheet.Shapes.AddChart2(0, chartType).Chart.SetSourceData dataRange
End Sub
