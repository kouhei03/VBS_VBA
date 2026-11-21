Attribute VB_Name = "Mod46_電話郵便番号"
Option Explicit

Public Sub 電話郵便番号のハイフン統一()
    Dim rng As Range
    Set rng = IIf(TypeName(Selection) = "Range", Selection, ActiveSheet.UsedRange)
    Dim arr As Variant: arr = rng.Value
    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If VarType(arr(r, c)) = vbString Then
                arr(r, c) = Replace(Replace(CStr(arr(r, c)), "－", "-"), "ー", "-")
            End If
        Next c
    Next r
    rng.Value = arr
End Sub
