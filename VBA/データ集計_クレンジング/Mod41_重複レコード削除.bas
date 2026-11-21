Attribute VB_Name = "Mod41_重複レコード削除"
Option Explicit
' ファイル処理概要: 辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 重複レコードを削除を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「重複レコードを削除」を選択して実行する。

' ── 重複レコードを削除 ──
Public Sub 重複レコードを削除()
    On Error GoTo Failed
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    If lastRow < 3 Then Exit Sub
    If MsgBox("重複レコードを削除しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Dim keyCol As String
    keyCol = InputBox("重複判定に使う列番号を入力してください", "重複レコード削除", "1")
    If Not IsNumeric(keyCol) Then Exit Sub
    SpeedUp True

    Dim col As Long: col = CLng(keyCol)
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim delRange As Range
    Dim r As Long
    For r = 2 To lastRow
        Dim key As String: key = CStr(ws.Cells(r, col).Value)
        If dic.Exists(key) Then
            If delRange Is Nothing Then Set delRange = ws.Rows(r) Else Set delRange = Union(delRange, ws.Rows(r))
        Else
            dic.Add key, r
        End If
    Next r
    If Not delRange Is Nothing Then delRange.Delete
CleanExit:
    SpeedUp False
    Exit Sub
Failed:
    SpeedUp False
    MsgBox "重複レコード削除でエラー: " & Err.Description, vbCritical
End Sub




