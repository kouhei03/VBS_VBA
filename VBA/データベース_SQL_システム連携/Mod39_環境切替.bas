Attribute VB_Name = "Mod39_環境切替"
Option Explicit

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
