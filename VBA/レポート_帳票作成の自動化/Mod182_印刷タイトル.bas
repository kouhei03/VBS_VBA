Attribute VB_Name = "Mod182_印刷タイトル"
Option Explicit

Public Sub 印刷タイトル行列固定設定()
    Dim repeatRows As String: repeatRows = InputBox("繰り返し行（例: $1:$2、空欄可）", "印刷タイトル")
    Dim repeatColumns As String: repeatColumns = InputBox("繰り返し列（例: $A:$B、空欄可）", "印刷タイトル")
    With ActiveSheet.PageSetup
        If repeatRows <> "" Then .PrintTitleRows = repeatRows
        If repeatColumns <> "" Then .PrintTitleColumns = repeatColumns
    End With
End Sub
