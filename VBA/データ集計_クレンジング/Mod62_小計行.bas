Attribute VB_Name = "Mod62_小計行"
Option Explicit

Public Sub 小計行を挿入()
    Dim ws As Worksheet: Set ws = ActiveSheet
    ws.UsedRange.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(3), _
                          Replace:=True, PageBreaks:=False, SummaryBelowData:=True
End Sub
