Attribute VB_Name = "Mod160_情シス定例サマリ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 情シス定例報告運用健全性サマリ作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「情シス定例報告運用健全性サマリ作成」を選択して実行する。

' ── 情シス定例報告運用健全性サマリ作成 ──
Public Sub 情シス定例報告運用健全性サマリ作成()
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:B1").Value = Array("指標", "件数")
    Dim r As Long: r = 2
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "定例サマリ" Then
            output.Cells(r, 1).Value = ws.Name
            output.Cells(r, 2).Value = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
            r = r + 1
        End If
    Next ws
    output.Columns("A:B").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("定例サマリ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "定例サマリ"
    End If
End Function




