Attribute VB_Name = "Mod40_完了通知"
Option Explicit

Public Sub 処理完了通知(Optional ByVal message As String = "処理が完了しました。")
    On Error Resume Next
    Beep
    On Error GoTo 0
    MsgBox message, vbInformation, "完了"
End Sub
