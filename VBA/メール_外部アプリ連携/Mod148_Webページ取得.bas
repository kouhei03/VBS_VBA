Attribute VB_Name = "Mod148_Webページ取得"
Option Explicit
' ファイル処理概要: HTTP通信を実行する。
' ファイルないコード処理詳細解説: Webページデータ取得シート転記を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、インターネットまたは社内ネットワーク、HTTP/MSXML利用可。
' ファイル実行例: ExcelのAlt+F8から「Webページデータ取得シート転記」を選択して実行する。

' ── Webページデータ取得シート転記 ──
Public Sub Webページデータ取得シート転記()
    Dim url As String: url = InputBox("取得するURL", "Webページ取得")
    If Trim$(url) = "" Then Exit Sub
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", url, False
    http.Send
    Dim html As Object: Set html = CreateObject("HTMLFile")
    html.write http.responseText
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "Web取得")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "URL"
    ws.Cells(1, 2).Value = url
    ws.Cells(2, 1).Value = "本文"
    ws.Cells(2, 2).Value = html.body.innerText
    ws.Cells(3, 1).Value = "状態"
    ws.Cells(3, 2).Value = http.Status
End Sub




