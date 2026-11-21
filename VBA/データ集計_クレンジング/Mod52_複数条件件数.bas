Attribute VB_Name = "Mod52_複数条件件数"
Option Explicit

Public Sub 複数条件で件数カウント()
    Dim col1 As Long: col1 = CLng(InputBox("条件列1", "件数カウント", "1"))
    Dim value1 As String: value1 = InputBox("条件値1", "件数カウント")
    Dim col2 As Long: col2 = CLng(InputBox("条件列2", "件数カウント", "2"))
    Dim value2 As String: value2 = InputBox("条件値2", "件数カウント")
    Dim count As Long
    Dim r As Long
    For r = 2 To GetLastRow(ActiveSheet)
        If CStr(ActiveSheet.Cells(r, col1).Value) = value1 And _
           CStr(ActiveSheet.Cells(r, col2).Value) = value2 Then count = count + 1
    Next r
    MsgBox "該当件数: " & count & " 件", vbInformation
End Sub
