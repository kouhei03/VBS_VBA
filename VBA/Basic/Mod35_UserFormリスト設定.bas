Attribute VB_Name = "Mod35_UserFormリスト設定"
Option Explicit

Public Sub コンボボックスを設定(ByVal combo As Object, ParamArray items() As Variant)
    Dim i As Long
    combo.Clear
    For i = LBound(items) To UBound(items)
        combo.AddItem CStr(items(i))
    Next i
    If combo.ListCount > 0 Then combo.ListIndex = 0
End Sub

Public Sub リストボックスを設定(ByVal listBox As Object, ByVal values As Variant)
    Dim i As Long
    listBox.Clear
    For i = LBound(values, 1) To UBound(values, 1)
        listBox.AddItem CStr(values(i, 1))
    Next i
End Sub
