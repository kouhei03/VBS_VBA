Attribute VB_Name = "Mod183_改ページ"
Option Explicit

Public Sub 手動改ページ挿入調整()
    If Selection Is Nothing Then Exit Sub
    ActiveSheet.HPageBreaks.Add Before:=Selection.Rows(1)
    MsgBox "選択行の前に改ページを挿入しました。", vbInformation
End Sub
