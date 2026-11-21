Attribute VB_Name = "Mod147_HTTPリクエスト"
Option Explicit

Public Sub WebAPIHTTPリクエストレスポンス取得()
    Dim url As String: url = InputBox("URLを入力", "HTTPリクエスト")
    If Trim$(url) = "" Then Exit Sub
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.SetRequestHeader "Accept", "application/json"
    http.Send
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "HTTPレスポンス")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "ステータス"
    ws.Cells(1, 2).Value = http.Status
    ws.Cells(2, 1).Value = "レスポンス"
    ws.Cells(2, 2).Value = http.responseText
End Sub
