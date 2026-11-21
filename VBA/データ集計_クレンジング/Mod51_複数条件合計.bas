Attribute VB_Name = "Mod51_複数条件合計"
Option Explicit

Public Sub 複数条件で合計()
    Dim rng As Range: Set rng = ActiveSheet.UsedRange
    Dim sumCol As Long: sumCol = CLng(InputBox("合計列番号", "複数条件合計", "3"))
    Dim criteriaCol As Long: criteriaCol = CLng(InputBox("条件列番号", "複数条件合計", "1"))
    Dim criterion As String: criterion = InputBox("条件値", "複数条件合計")
    Dim total As Double
    Dim r As Long
    For r = 2 To GetLastRow(ActiveSheet)
        If CStr(ActiveSheet.Cells(r, criteriaCol).Value) = criterion Then
            If IsNumeric(ActiveSheet.Cells(r, sumCol).Value) Then
                total = total + CDbl(ActiveSheet.Cells(r, sumCol).Value)
            End If
        End If
    Next r
    MsgBox "合計: " & Format$(total, "#,##0"), vbInformation
End Sub
