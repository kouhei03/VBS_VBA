Attribute VB_Name = "Mod170_表示形式"
Option Explicit

Public Sub 数値表示形式設定()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim mode As String
    mode = InputBox("形式を入力: 通貨 / 桁区切り / パーセント / 日付", "表示形式")
    Dim formatText As String
    Select Case mode
        Case "通貨": formatText = "[$￥-ja-JP]#,##0"
        Case "桁区切り": formatText = "#,##0"
        Case "パーセント": formatText = "0.0%"
        Case "日付": formatText = "yyyy/mm/dd"
        Case Else: Exit Sub
    End Select
    If Not Selection Is Nothing Then Selection.NumberFormat = formatText
End Sub
