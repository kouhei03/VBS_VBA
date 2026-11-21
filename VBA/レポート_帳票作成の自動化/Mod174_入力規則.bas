Attribute VB_Name = "Mod174_入力規則"
Option Explicit

Public Sub 入力規則ドロップダウン数値範囲設定()
    If Selection Is Nothing Then Exit Sub
    Dim mode As String
    mode = InputBox("設定を入力: 1=リスト、2=0以上の数値", "入力規則")
    Selection.Validation.Delete
    If mode = "1" Then
        Dim listText As String
        listText = InputBox("カンマ区切りリストを入力", "入力規則")
        If listText <> "" Then Selection.Validation.Add xlValidateList, , , listText
    ElseIf mode = "2" Then
        Selection.Validation.Add xlValidateDecimal, xlValidAlertStop, xlBetween, "0", "100000000"
    End If
End Sub
