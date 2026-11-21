Attribute VB_Name = "Mod134_タスク登録"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: OutlookタスクToDo登録を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「OutlookタスクToDo登録」を選択して実行する。

' ── OutlookタスクToDo登録 ──
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




