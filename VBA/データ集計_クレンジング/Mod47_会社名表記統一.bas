Attribute VB_Name = "Mod47_会社名表記統一"
Option Explicit

Public Sub 会社名表記を統一()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim arr As Variant: arr = Selection.Value
    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If VarType(arr(r, c)) = vbString Then
                arr(r, c) = NormalizeCompanyName(CStr(arr(r, c)))
            End If
        Next c
    Next r
    Selection.Value = arr
End Sub
