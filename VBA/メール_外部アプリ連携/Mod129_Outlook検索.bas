Attribute VB_Name = "Mod129_Outlook検索"
Option Explicit

Public Sub Outlook受信トレイ送信済みアイテム検索()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim namespace As Object: Set namespace = outlook.GetNamespace("MAPI")
    Dim keyword As String: keyword = InputBox("検索キーワード", "Outlook検索")
    If Trim$(keyword) = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "Outlook検索")
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("フォルダ", "件名", "受信日時")
    Dim r As Long: r = 2
    SearchFolder namespace.GetDefaultFolder(6), keyword, ws, r
    SearchFolder namespace.GetDefaultFolder(5), keyword, ws, r
    ws.Columns("A:C").AutoFit
End Sub

Private Sub SearchFolder(ByVal folder As Object, ByVal keyword As String, ByVal ws As Worksheet, ByRef r As Long)
    Dim item As Object
    For Each item In folder.Items
        If InStr(1, CStr(item.Subject), keyword, vbTextCompare) > 0 Then
            ws.Cells(r, 1).Value = folder.Name
            ws.Cells(r, 2).Value = item.Subject
            ws.Cells(r, 3).Value = item.ReceivedTime
            r = r + 1
        End If
    Next item
End Sub
