Attribute VB_Name = "Mod134_タスク登録"
Option Explicit

Public Sub OutlookタスクToDo登録()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim task As Object: Set task = outlook.CreateItem(3)
    task.Subject = InputBox("タスク名", "タスク登録")
    task.Body = InputBox("詳細", "タスク登録")
    Dim dueDate As String: dueDate = InputBox("期限（例: 2026/12/01、空欄可）", "タスク登録")
    If Trim$(dueDate) <> "" Then
        task.DueDate = CDate(dueDate)
        task.ReminderSet = True
    End If
    task.Save
    MsgBox "タスクを登録しました。", vbInformation
End Sub
