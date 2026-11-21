Attribute VB_Name = "Mod189_帳票印刷"
Option Explicit

Public Sub 帳票印刷プリンタ部数指定()
    Dim printerName As String: printerName = InputBox("プリンタ名（空欄で既定）", "帳票印刷")
    Dim copies As Variant
    copies = Application.InputBox("印刷部数", "帳票印刷", 1, , , , , 1)
    If copies = False Then Exit Sub
    If Trim$(printerName) <> "" Then ActiveSheet.Application.ActivePrinter = printerName
    ActiveSheet.PrintOut Copies:=CLng(copies), Collate:=True
End Sub
