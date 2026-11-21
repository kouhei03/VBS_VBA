Attribute VB_Name = "Mod114_フェーズ移行"
Option Explicit

Public Sub 案件フェーズ移行チェック項目達成確認()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim complete As Boolean: complete = True
    Dim r As Long
    For r = 2 To lastRow
        If UCase$(CStr(ws.Cells(r, 2).Value)) <> "完了" Then complete = False
    Next r
    ws.Range("E1").Value = "フェーズ移行判定"
    ws.Range("E2").Value = IIf(complete, "移行可", "移行不可")
End Sub
