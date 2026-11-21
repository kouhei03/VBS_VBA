Attribute VB_Name = "Mod151_システム情報"
Option Explicit

Public Sub Windowsシステム情報環境変数取得()
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim environment As Object: Set environment = shell.Environment("PROCESS")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "システム情報")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "項目"
    ws.Cells(1, 2).Value = "値"
    ws.Cells(2, 1).Value = "COMPUTERNAME"
    ws.Cells(2, 2).Value = environment("COMPUTERNAME")
    ws.Cells(3, 1).Value = "USERNAME"
    ws.Cells(3, 2).Value = environment("USERNAME")
    ws.Cells(4, 1).Value = "SystemRoot"
    ws.Cells(4, 2).Value = environment("SystemRoot")
    ws.Cells(5, 1).Value = "TEMP"
    ws.Cells(5, 2).Value = environment("TEMP")
    ws.Columns("A:B").AutoFit
End Sub
