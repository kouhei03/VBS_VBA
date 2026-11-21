Attribute VB_Name = "Mod109_チケット集計"
Option Explicit

Public Sub チケット対応状況平均解決時間集計()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim count As Long, totalHours As Double
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 4).Value) = "完了" _
           And IsDate(ws.Cells(r, 2).Value) And IsDate(ws.Cells(r, 3).Value) Then
            count = count + 1
            totalHours = totalHours + CDbl(CDate(ws.Cells(r, 3).Value) - CDate(ws.Cells(r, 2).Value)) * 24
        End If
    Next r
    Dim avg As Double: avg = 0
    If count > 0 Then avg = totalHours / count
    MsgBox "完了件数: " & count & vbCrLf & "平均解決時間: " & Format$(avg, "0.0") & " 時間", vbInformation
End Sub
