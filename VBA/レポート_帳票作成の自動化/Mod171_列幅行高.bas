Attribute VB_Name = "Mod171_列幅行高"
Option Explicit

Public Sub 列幅行高自動調整余白確保()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Columns.AutoFit
    ws.UsedRange.Rows.AutoFit
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count
        If ws.Cells(1, c).EntireColumn.ColumnWidth > 30 Then
            ws.Cells(1, c).EntireColumn.ColumnWidth = 30
        End If
    Next c
    MsgBox "列幅と行高さを調整しました。", vbInformation
End Sub
