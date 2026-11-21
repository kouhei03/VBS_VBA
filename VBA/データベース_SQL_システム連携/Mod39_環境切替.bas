Attribute VB_Name = "Mod39_環境切替"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テスト本番接続先切替制御を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テスト本番接続先切替制御」を選択して実行する。

' ── テスト本番接続先切替制御 ──
Public Sub テスト本番接続先切替制御()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("接続設定")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "接続設定シートを作成してください。", vbExclamation
        Exit Sub
    End If
    Dim environment As String
    environment = InputBox("利用環境を入力: テスト / 本番", "環境切替")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CStr(ws.Cells(r, 1).Value) = environment Then
            ThisWorkbook.Names.Add Name:="ActiveConnectionString", RefersTo:="='" & ws.Name & "'!$B$" & r
            MsgBox environment & "環境を選択しました。", vbInformation
            Exit Sub
        End If
    Next r
    MsgBox "指定された環境が見つかりません。", vbExclamation
End Sub




