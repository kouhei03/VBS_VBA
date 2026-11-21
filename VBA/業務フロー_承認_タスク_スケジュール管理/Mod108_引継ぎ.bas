Attribute VB_Name = "Mod108_引継ぎ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 業務引継ぎ引継チェックリスト進捗管理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「業務引継ぎ引継チェックリスト進捗管理」を選択して実行する。

' ── 業務引継ぎ引継チェックリスト進捗管理 ──
Public Sub 業務引継ぎ引継チェックリスト進捗管理()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("引継項目", "担当者", "期限", "完了")
    Dim r As Long: r = 2
    For r = 2 To ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        ActiveSheet.Rows(r).Copy ws.Rows(r)
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("引継チェックリスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "引継チェックリスト"
    End If
End Function




