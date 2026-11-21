Attribute VB_Name = "Mod79_式を値へ変換"
Option Explicit

Public Sub 計算式を値へ一括変換()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub
