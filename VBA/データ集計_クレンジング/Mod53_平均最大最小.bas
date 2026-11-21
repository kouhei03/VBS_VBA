Attribute VB_Name = "Mod53_平均最大最小"
Option Explicit

Public Sub 平均最大最小を算出()
    Dim col As Long
    col = CLng(InputBox("数値列番号を入力", "統計算出", "3"))
    Dim values As Collection: Set values = New Collection
    Dim r As Long
    For r = 2 To GetLastRow(ActiveSheet)
        If IsNumeric(ActiveSheet.Cells(r, col).Value) Then
            values.Add CDbl(ActiveSheet.Cells(r, col).Value)
        End If
    Next r
    If values.Count = 0 Then
        MsgBox "数値データがありません。", vbExclamation
        Exit Sub
    End If
    Dim sum As Double, minimum As Double, maximum As Double
    minimum = values(1): maximum = values(1)
    Dim i As Long
    For i = 1 To values.Count
        sum = sum + values(i)
        If values(i) < minimum Then minimum = values(i)
        If values(i) > maximum Then maximum = values(i)
    Next i
    MsgBox "件数: " & values.Count & vbCrLf & _
           "平均: " & Format$(sum / values.Count, "#,##0.00") & vbCrLf & _
           "最大: " & Format$(maximum, "#,##0") & vbCrLf & _
           "最小: " & Format$(minimum, "#,##0"), vbInformation
End Sub
