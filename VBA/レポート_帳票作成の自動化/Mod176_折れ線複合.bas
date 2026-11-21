Attribute VB_Name = "Mod176_折れ線複合"
Option Explicit

Public Sub 折れ線複合グラフ自動生成()
    Dim dataRange As Range
    Set dataRange = Application.InputBox("グラフ化する範囲を選択", "折れ線グラフ", Selection.Address, , , , , 8)
    Dim chart As Chart
    Set chart = ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Chart
    chart.SetSourceData dataRange
    chart.HasTitle = True
    chart.ChartTitle.Text = "推移グラフ"
End Sub
