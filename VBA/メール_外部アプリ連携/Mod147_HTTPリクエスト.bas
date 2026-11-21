Attribute VB_Name = "Mod147_HTTPリクエスト"
Option Explicit
' ファイル処理概要: HTTP通信を実行する。
' ファイルないコード処理詳細解説: WebAPIHTTPリクエストレスポンス取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、インターネットまたは社内ネットワーク、HTTP/MSXML利用可。
' ファイル実行例: ExcelのAlt+F8から「WebAPIHTTPリクエストレスポンス取得」を選択して実行する。

' ── WebAPIHTTPリクエストレスポンス取得 ──
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




