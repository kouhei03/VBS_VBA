Attribute VB_Name = "Mod21_文字列分割"
Option Explicit

Public Sub 区切り文字で分割()
    Dim text As String
    text = InputBox("分割する文字列を入力してください", "文字列分割", "東京,大阪,福岡")
    If StrPtr(text) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(text, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Debug.Print i & ": " & parts(i)
    Next i
End Sub
