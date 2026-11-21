Attribute VB_Name = "Mod130_メール条件抽出"
Option Explicit

Public Sub 指定条件メール件名送信元受信日時抽出()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim namespace As Object: Set namespace = outlook.GetNamespace("MAPI")
    Dim inbox As Object: Set inbox = namespace.GetDefaultFolder(6)
    Dim senderFilter As String: senderFilter = InputBox("送信元に含む文字", "メール抽出")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "メール抽出")
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("件名", "送信元", "受信日時", "本文先頭")
    Dim r As Long: r = 2
    Dim item As Object
    For Each item In inbox.Items
        If InStr(1, CStr(item.SenderName), senderFilter, vbTextCompare) > 0 _
           Or Trim$(senderFilter) = "" Then
            ws.Cells(r, 1).Value = item.Subject
            ws.Cells(r, 2).Value = item.SenderName
            ws.Cells(r, 3).Value = item.ReceivedTime
            ws.Cells(r, 4).Value = Left$(item.Body, 200)
            r = r + 1
        End If
    Next item
    ws.Columns("A:D").AutoFit
End Sub
