Attribute VB_Name = "Mod67_桁数検出"
Option Explicit

Public Sub 桁数不足超過を検出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = CLng(InputBox("対象列番号", "桁数検出", "1"))
    Dim minLen As Long: minLen = CLng(InputBox("最小桁数", "桁数検出", "1"))
    Dim maxLen As Long: maxLen = CLng(InputBox("最大桁数", "桁数検出", "10"))
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        Dim length As Long: length = Len(CStr(ws.Cells(r, col).Value))
        If length < minLen Or length > maxLen Then
            ws.Cells(r, col).Interior.Color = RGB(255, 192, 0)
        End If
    Next r
End Sub
