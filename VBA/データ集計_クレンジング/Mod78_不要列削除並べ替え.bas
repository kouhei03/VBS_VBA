Attribute VB_Name = "Mod78_不要列削除並べ替え"
Option Explicit

Public Sub 不要列を削除()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim colText As String: colText = InputBox("削除する列番号（カンマ区切り）", "列削除")
    If colText = "" Then Exit Sub
    Dim parts As Variant: parts = Split(colText, ",")
    Dim i As Long
    For i = UBound(parts) To LBound(parts) Step -1
        ws.Columns(CLng(Trim$(CStr(parts(i))))).Delete
    Next i
End Sub
