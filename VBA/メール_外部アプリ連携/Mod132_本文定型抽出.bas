Attribute VB_Name = "Mod132_本文定型抽出"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 受信本文定型データ抽出シート転記を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「受信本文定型データ抽出シート転記」を選択して実行する。

' ── 受信本文定型データ抽出シート転記 ──
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




