Attribute VB_Name = "Mod193_担当者専用シート"
Option Explicit
' ファイル処理概要: ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: 担当者案件専用作業シート自動生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「担当者案件専用作業シート自動生成」を選択して実行する。

' ── 担当者案件専用作業シート自動生成 ──
Public Sub 担当者案件専用作業シート自動生成()
    Dim owner As String: owner = InputBox("担当者名または案件名", "専用シート")
    If Trim$(owner) = "" Then Exit Sub
    Dim name As String: name = Replace(owner, "/", "_")
    If Len(name) > 31 Then name = Left$(name, 31)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = name
    ws.Range("A1").Value = "担当者/案件"
    ws.Range("B1").Value = owner
End Sub




