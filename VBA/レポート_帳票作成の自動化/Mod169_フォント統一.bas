Attribute VB_Name = "Mod169_フォント統一"
Option Explicit

Public Sub フォント種類サイズ太字文字色統一()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Font.Name = "Yu Gothic"
    ws.UsedRange.Font.Size = 10
    ws.UsedRange.Font.Color = RGB(32, 32, 32)
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
    ws.Rows(1).Interior.Color = RGB(47, 85, 151)
End Sub
