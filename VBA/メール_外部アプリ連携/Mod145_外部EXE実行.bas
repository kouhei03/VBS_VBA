Attribute VB_Name = "Mod145_外部EXE実行"
Option Explicit

Public Sub 外部EXEバッチスクリプト実行()
    Dim command As String
    command = InputBox("実行するコマンドまたはフルパスを入力", "外部プログラム実行")
    If Trim$(command) = "" Then Exit Sub
    CreateObject("WScript.Shell").Run command, 1, False
End Sub
