Attribute VB_Name = "Mod154_外部アプリ待機"
Option Explicit

Public Sub 外部アプリ起動時プロセス待機制御()
    Dim exePath As String: exePath = InputBox("起動するEXEのフルパス", "プロセス待機")
    If Trim$(exePath) = "" Then Exit Sub
    CreateObject("WScript.Shell").Run """" & exePath & """", 1, True
    MsgBox "プロセスが終了しました。", vbInformation
End Sub
