Attribute VB_Name = "Mod08_エラー処理"
Option Explicit

Public Sub 安全な処理例()
    On Error GoTo Failed

    ' ここに処理を記述
    Dim value As Long
    value = CLng(ActiveSheet.Range("A1").Value)

CleanExit:
    Exit Sub
Failed:
    MsgBox "エラー " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub
