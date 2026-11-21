Attribute VB_Name = "Mod155_連携エラー切分"
Option Explicit

Public Sub 複数アプリ連携エラー切り分けクローズ処理()
    On Error GoTo Failed
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim word As Object: Set word = CreateObject("Word.Application")
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    MsgBox "Outlook、Word、PowerPointへの接続に成功しました。", vbInformation
CleanExit:
    On Error Resume Next
    If Not word Is Nothing Then word.Quit
    If Not ppt Is Nothing Then ppt.Quit
    Exit Sub
Failed:
    MsgBox "連携エラー: " & Err.Description & " (番号 " & Err.Number & ")", vbCritical
    Resume CleanExit
End Sub
