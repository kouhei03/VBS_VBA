Attribute VB_Name = "Mod194_勤怠シフト"
Option Explicit

Public Sub 勤怠集計シフト表整形出力()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Columns.AutoFit
    ws.UsedRange.Rows.AutoFit
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(47, 85, 151)
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 0
End Sub
