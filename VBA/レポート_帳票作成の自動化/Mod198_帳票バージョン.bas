Attribute VB_Name = "Mod198_帳票バージョン"
Option Explicit

Public Sub 帳票バージョン作成日時自動印字()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim version As String: version = InputBox("バージョン（例: v1.0）", "帳票バージョン", "v1.0")
    With ws.PageSetup
        .RightHeader = "Ver. " & version
        .RightFooter = "作成日時: " & Format$(Now, "yyyy/mm/dd hh:mm")
    End With
    ws.Cells(1, ws.UsedRange.Columns.Count + 1).Value = "Ver." & version
    ws.Cells(2, ws.UsedRange.Columns.Count).Value = "作成: " & Format$(Now, "yyyy/mm/dd hh:mm")
End Sub
