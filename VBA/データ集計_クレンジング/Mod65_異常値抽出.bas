Attribute VB_Name = "Mod65_異常値抽出"
Option Explicit

Public Sub 異常値を抽出してフラグ付与()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = CLng(InputBox("対象列番号", "異常値検出", "3"))
    Dim threshold As Double: threshold = CDbl(InputBox("閾値", "異常値検出", "100"))
    Dim lastCol As Long: lastCol = Application.WorksheetFunction.Max(GetLastCol(ws), col + 1)
    ws.Cells(1, lastCol).Value = "異常値フラグ"
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        If IsNumeric(ws.Cells(r, col).Value) Then
            ws.Cells(r, lastCol).Value = IIf(CDbl(ws.Cells(r, col).Value) > threshold, "異常", "")
            If CDbl(ws.Cells(r, col).Value) > threshold Then
                ws.Cells(r, col).Interior.Color = RGB(255, 199, 206)
            End If
        End If
    Next r
End Sub
