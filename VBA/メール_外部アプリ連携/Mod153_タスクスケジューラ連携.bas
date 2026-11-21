Attribute VB_Name = "Mod153_タスクスケジューラ連携"
Option Explicit

Public Sub タスクスケジューラ連携起動設計()
    Dim scriptPath As String: scriptPath = InputBox("定期実行するVBSまたはBATのパス", "タスクスケジューラ")
    Dim taskName As String: taskName = InputBox("タスク名", "タスクスケジューラ")
    If Trim$(scriptPath) = "" Or Trim$(taskName) = "" Then Exit Sub
    Dim command As String
    command = "schtasks.exe /Create /TN """ & taskName & """ /TR """"" & scriptPath & """" & """ /SC DAILY /ST 09:00"
    MsgBox "以下のschtasksコマンドを管理者権限で実行してください。" & vbCrLf & command, vbInformation
End Sub
