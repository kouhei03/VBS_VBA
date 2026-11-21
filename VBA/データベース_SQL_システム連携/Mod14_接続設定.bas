Attribute VB_Name = "Mod14_接続設定"
Option Explicit

Public Sub 接続文字列外部設定シート管理()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells(1, 1).Value = "環境"
    ws.Cells(1, 2).Value = "接続文字列"
    ws.Cells(2, 1).Value = "テスト"
    ws.Cells(3, 1).Value = "本番"
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("接続設定")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "接続設定"
    End If
End Function
