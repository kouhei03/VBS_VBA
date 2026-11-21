Attribute VB_Name = "Mod118_申請番号採番"
Option Explicit

Public Sub 申請管理番号自動採番重複防止()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim prefix As String: prefix = InputBox("番号の接頭辞", "番号採番", "REQ")
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then
            Dim candidate As String
            candidate = prefix & "-" & Format$(Now, "yyyymmdd") & "-" & Format$(r - 1, "000")
            ws.Cells(r, 1).Value = candidate
            Exit Sub
        End If
    Next r
End Sub
