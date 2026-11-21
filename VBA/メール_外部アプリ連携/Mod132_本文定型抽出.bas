Attribute VB_Name = "Mod132_本文定型抽出"
Option Explicit

Public Sub 受信本文定型データ抽出シート転記()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim inbox As Object: Set inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "受信本文抽出")
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("件名", "受信日時", "送信元", "本文")
    Dim r As Long: r = 2
    Dim item As Object
    For Each item In inbox.Items
        ws.Cells(r, 1).Value = item.Subject
        ws.Cells(r, 2).Value = item.ReceivedTime
        ws.Cells(r, 3).Value = item.SenderName
        ws.Cells(r, 4).Value = item.Body
        r = r + 1
    Next item
    ws.Columns("A:D").AutoFit
End Sub
