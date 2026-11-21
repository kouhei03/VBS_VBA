Attribute VB_Name = "Mod61_並び替え"
Option Explicit

Public Sub 複数キーで並び替え()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim keyText As String: keyText = InputBox("並び替え列番号（例: 1,3）", "並び替え")
    If keyText = "" Then Exit Sub
    Dim parts As Variant: parts = Split(keyText, ",")
    If UBound(parts) = 0 Then
        ws.UsedRange.Sort Key1:=ws.Cells(1, CLng(parts(0))), Header:=xlYes
    Else
        ws.UsedRange.Sort Key1:=ws.Cells(1, CLng(parts(0))), _
                          Key2:=ws.Cells(1, CLng(parts(1))), Header:=xlYes
    End If
End Sub
