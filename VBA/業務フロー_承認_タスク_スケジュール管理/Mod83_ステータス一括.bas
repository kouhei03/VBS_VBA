Attribute VB_Name = "Mod83_ステータス一括"
Option Explicit

Public Sub ステータス一括更新()
    Dim status As String
    status = InputBox("更新ステータス（未着手 / 進行中 / 完了 / 保留）", "ステータス更新")
    If Selection Is Nothing Then Exit Sub
    Selection.Value = status
End Sub
