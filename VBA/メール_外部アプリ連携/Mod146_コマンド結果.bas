Attribute VB_Name = "Mod146_コマンド結果"
Option Explicit

Public Sub コマンドプロンプト実行結果取得()
    Dim command As String: command = InputBox("実行するコマンドを入力", "コマンド実行")
    If Trim$(command) = "" Then Exit Sub
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim output As String
    output = shell.Exec("%COMSPEC% /c " & command).StdOut.ReadAll
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "コマンド結果")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = output
End Sub
