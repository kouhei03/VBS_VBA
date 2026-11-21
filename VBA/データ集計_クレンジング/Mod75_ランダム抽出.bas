Attribute VB_Name = "Mod75_ランダム抽出"
Option Explicit

Public Sub ランダムサンプリング抽出()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim sampleSize As Long: sampleSize = CLng(InputBox("抽出件数", "ランダム抽出", "10"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "ランダムサンプル")
    out.Cells.Clear
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    out.Range(out.Cells(1, 1), out.Cells(1, lastCol)).Value = _
        ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Randomize
    Dim outputRow As Long: outputRow = 2
    Do While outputRow < sampleSize + 2 And dic.Count < lastRow - 1
        Dim r As Long: r = Int((lastRow - 1) * Rnd) + 2
        If Not dic.Exists(r) Then
            dic.Add r, True
            out.Rows(outputRow).Value = ws.Rows(r).Value
            outputRow = outputRow + 1
        End If
    Loop
End Sub
