Attribute VB_Name = "Mod68_重複キー一覧"
Option Explicit

Public Sub 重複キー一覧と件数を表示()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = CLng(InputBox("キー列番号", "重複キー一覧", "1"))
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To GetLastRow(ws)
        Dim key As String: key = CStr(ws.Cells(r, col).Value)
        If Len(key) > 0 Then
            If Not dic.Exists(key) Then dic.Add key, 0
            dic(key) = dic(key) + 1
        End If
    Next r
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "重複キー一覧")
    out.Cells.Clear
    out.Cells(1, 1).Value = "キー": out.Cells(1, 2).Value = "件数"
    r = 2
    Dim k As Variant
    For Each k In dic.Keys
        If dic(k) > 1 Then
            out.Cells(r, 1).Value = k
            out.Cells(r, 2).Value = dic(k)
            r = r + 1
        End If
    Next k
End Sub
