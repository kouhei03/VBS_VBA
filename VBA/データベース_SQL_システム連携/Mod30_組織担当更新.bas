Attribute VB_Name = "Mod30_組織担当更新"
Option Explicit

Public Sub 組織部署マスタ変更履歴担当割当更新()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        If CStr(ws.Cells(r, 4).Value) = "変更" Then
            ws.Cells(r, 5).Value = "担当割当更新対象"
        Else
            ws.Cells(r, 5).Value = ""
        End If
    Next r
End Sub
