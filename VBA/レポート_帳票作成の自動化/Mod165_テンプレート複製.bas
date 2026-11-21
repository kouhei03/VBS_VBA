Attribute VB_Name = "Mod165_テンプレート複製"
Option Explicit

Public Sub テンプレートシート複製対象名リネーム()
    Dim source As Worksheet: Set source = ActiveSheet
    Dim targetName As String: targetName = InputBox("複製後のシート名", "テンプレート複製")
    If Trim$(targetName) = "" Then Exit Sub
    Dim copy As Worksheet
    source.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set copy = ActiveSheet
    copy.Name = targetName
    copy.Cells.Replace What:="{対象名}", Replacement:=targetName, LookAt:=xlPart
    MsgBox "テンプレートを複製しました。", vbInformation
End Sub
