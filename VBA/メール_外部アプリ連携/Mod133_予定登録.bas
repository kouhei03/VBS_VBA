Attribute VB_Name = "Mod133_予定登録"
Option Explicit

Public Sub Outlook予定表予定登録()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim appointment As Object: Set appointment = outlook.CreateItem(1)
    appointment.Subject = InputBox("予定の件名", "予定登録")
    appointment.Start = CDate(InputBox("開始日時（例: 2026/12/01 10:00）", "予定登録"))
    appointment.End = CDate(InputBox("終了日時（例: 2026/12/01 11:00）", "予定登録"))
    appointment.Location = InputBox("場所", "予定登録")
    appointment.Body = InputBox("内容", "予定登録")
    appointment.ReminderSet = True
    appointment.Save
    MsgBox "予定を登録しました。", vbInformation
End Sub
