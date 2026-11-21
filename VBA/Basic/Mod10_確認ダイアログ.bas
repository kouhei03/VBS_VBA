Attribute VB_Name = "Mod10_確認ダイアログ"
Option Explicit

Public Function ユーザー確認(ByVal message As String) As Boolean
    ユーザー確認 = (MsgBox(message, vbQuestion + vbYesNo + vbDefaultButton2, _
                           "確認") = vbYes)
End Function
