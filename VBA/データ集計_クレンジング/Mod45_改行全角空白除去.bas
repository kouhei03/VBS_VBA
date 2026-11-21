Attribute VB_Name = "Mod45_改行全角空白除去"
Option Explicit

Public Sub 改行コードと全角空白を除去()
    Dim rng As Range
    Set rng = IIf(TypeName(Selection) = "Range", Selection, ActiveSheet.UsedRange)
    Dim arr As Variant: arr = rng.Value
    Dim r As Long, c As Long
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If VarType(arr(r, c)) = vbString Then
                arr(r, c) = Replace(Replace(Replace(arr(r, c), Chr$(13), ""), Chr$(10), ""), "　", "")
            End If
        Next c
    Next r
    rng.Value = arr
End Sub
