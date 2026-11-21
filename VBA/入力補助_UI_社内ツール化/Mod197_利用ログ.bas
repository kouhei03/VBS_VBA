Attribute VB_Name = "Mod197_利用ログ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 機能別利用状況簡易ログ記録を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「機能別利用状況簡易ログ記録」を選択して実行する。

' ── 機能別利用状況簡易ログ記録 ──
Public Sub 機能別利用状況簡易ログ記録()
    Dim functionName As String: functionName = InputBox("機能名", "利用ログ")
    If Trim$(functionName) = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = GetOrCreateSheet
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = CreateObject("WScript.Network").UserName
    ws.Cells(r, 3).Value = functionName
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("利用ログ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "利用ログ"
    End If
End Function




