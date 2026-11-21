Attribute VB_Name = "Mod14_接続設定"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 接続文字列外部設定シート管理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「接続文字列外部設定シート管理」を選択して実行する。

' ── 接続文字列外部設定シート管理 ──
Public Sub 接続文字列外部設定シート管理()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells(1, 1).Value = "環境"
    ws.Cells(1, 2).Value = "接続文字列"
    ws.Cells(2, 1).Value = "テスト"
    ws.Cells(3, 1).Value = "本番"
    ws.Columns("A:B").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("接続設定")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "接続設定"
    End If
End Function




