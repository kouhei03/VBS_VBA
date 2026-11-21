Attribute VB_Name = "Mod22_文字列結合"
Option Explicit

Public Function 文字列を結合(ByVal separator As String, ParamArray values() As Variant) As String
    Dim result As String
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Len(CStr(values(i))) > 0 Then
            If Len(result) > 0 Then result = result & separator
            result = result & CStr(values(i))
        End If
    Next i
    文字列を結合 = result
End Function
