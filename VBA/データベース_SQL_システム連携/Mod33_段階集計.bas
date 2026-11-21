Attribute VB_Name = "Mod33_段階集計"
Option Explicit

Public Sub ビュー一時テーブル段階集計()
    Dim connectionString As String: connectionString = InputBox("接続文字列", "段階集計")
    Dim sql As String: sql = InputBox("一時テーブルまたはビューを作成するSQL", "段階集計")
    If connectionString = "" Or sql = "" Then Exit Sub
    Dim connection As Object: Set connection = CreateObject("ADODB.Connection")
    connection.Open connectionString
    connection.Execute sql
    connection.Close
    MsgBox "段階集計処理を実行しました。", vbInformation
End Sub
