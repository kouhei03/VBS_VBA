Attribute VB_Name = "Mod178_グラフ整形"
Option Explicit

Public Sub グラフ位置サイズタイトル凡例自動整形()
    Dim titleText As String: titleText = InputBox("グラフタイトル", "グラフ整形")
    Dim chartObject As ChartObject
    For Each chartObject In ActiveSheet.ChartObjects
        With chartObject
            .Left = 50
            .Top = 120
            .Width = 600
            .Height = 360
            .Chart.HasTitle = True
            .Chart.ChartTitle.Text = titleText
            .Chart.HasLegend = True
            .Chart.Legend.Position = xlLegendPositionBottom
        End With
    Next chartObject
End Sub
