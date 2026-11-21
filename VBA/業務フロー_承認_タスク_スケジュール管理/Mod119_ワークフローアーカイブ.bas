Attribute VB_Name = "Mod119_ワークフローアーカイブ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ワークフロー履歴アーカイブ世代管理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ワークフロー履歴アーカイブ世代管理」を選択して実行する。

' ── ワークフロー履歴アーカイブ世代管理 ──
Public Sub ワークフロー履歴アーカイブ世代管理()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ワークフロー履歴")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "ワークフロー履歴シートがありません。", vbExclamation
        Exit Sub
    End If
    Dim archive As Worksheet: Set archive = GetOrCreateSheet
    ws.UsedRange.Copy archive.Range("A1")
    ws.Cells.Clear
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    Dim name As String: name = "履歴_" & Format$(Now, "yyyymmdd_hhnnss")
    Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
    GetOrCreateSheet.Name = name
End Function




