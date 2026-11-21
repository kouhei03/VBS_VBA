Attribute VB_Name = "Mod106_アクション期限"
Option Explicit

Public Sub アクション期限管理次回会議リスト生成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("アクション", "担当者", "期限", "状態")
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 4).Value) <> "完了" Then
            ws.Rows(r).Copy output.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("次回会議リスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "次回会議リスト"
    End If
End Function
