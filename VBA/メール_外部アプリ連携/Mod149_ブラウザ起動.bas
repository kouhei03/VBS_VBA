Attribute VB_Name = "Mod149_ブラウザ起動"
Option Explicit

Public Sub 指定URLブラウザ起動()
    Dim url As String: url = InputBox("開くURL", "ブラウザ起動")
    If Trim$(url) = "" Then Exit Sub
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    shell.Run "explorer.exe """ & url & """", 1, False
End Sub
