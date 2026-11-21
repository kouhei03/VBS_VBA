Attribute VB_Name = "Mod191_表紙目次"
Option Explicit

Public Sub 表紙目次シート自動生成ハイパーリンク()
    Dim indexSheet As Worksheet: Set indexSheet = GetOrCreateSheet("目次")
    indexSheet.Cells.Clear
    indexSheet.Cells(1, 1).Value = "目次"
    indexSheet.Cells(1, 1).Font.Size = 18
    indexSheet.Cells(1, 1).Font.Bold = True
    Dim r As Long: r = 3
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "目次" Then
            indexSheet.Cells(r, 1).Value = ws.Name
            indexSheet.Hyperlinks.Add Anchor:=indexSheet.Cells(r, 1), Address:="", SubAddress:="'" & ws.Name & "'!A1"
            r = r + 1
        End If
    Next ws
    indexSheet.Columns("A:A").AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        GetOrCreateSheet.Name = name
    End If
End Function
