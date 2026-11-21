Attribute VB_Name = "Mod240_テスト完了レポート"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テスト完了レポート関係者共有サマリ生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テスト完了レポート関係者共有サマリ生成」を選択して実行する。

' ── テスト完了レポート関係者共有サマリ生成 ──
Public Sub テスト完了レポート関係者共有サマリ生成()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("項目", "内容")
    ws.Cells(2, 1).Value = "完了日時"
    ws.Cells(2, 2).Value = Now
    ws.Cells(3, 1).Value = "対象ブック"
    ws.Cells(3, 2).Value = ThisWorkbook.Name
    ws.Cells(4, 1).Value = "総評"
    ws.Cells(4, 2).Value = InputBox("リリース判定または総評を入力", "テスト完了レポート")
    ws.Columns("A:B").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("テスト完了レポート")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "テスト完了レポート"
    End If
End Function




