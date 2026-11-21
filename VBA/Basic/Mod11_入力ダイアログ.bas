Attribute VB_Name = "Mod11_入力ダイアログ"
Option Explicit

Public Function 入力値を取得(ByVal prompt As String, Optional ByVal defaultValue As String = "") As Variant
    Dim result As String
    result = InputBox(prompt, "入力", defaultValue)
    If StrPtr(result) = 0 Or Trim$(result) = "" Then
        入力値を取得 = Empty
    Else
        入力値を取得 = result
    End If
End Function
