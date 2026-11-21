Attribute VB_Name = "Mod118_ハッシュ差分"
Option Explicit

Public Sub ファイルハッシュ更新チェック差分処理()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    Dim hash As String: hash = GetSha256(filePath)
    If hash = "" Then
        MsgBox "ハッシュを取得できませんでした。", vbCritical
        Exit Sub
    End If
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("ハッシュ差分")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "ファイル"
    ws.Cells(1, 2).Value = "SHA-256"
    ws.Cells(2, 1).Value = filePath
    ws.Cells(2, 2).Value = hash
End Sub

Private Function GetSha256(ByVal filePath As String) As String
    On Error GoTo Failed
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    Dim output As String
    output = shell.Exec("certutil.exe -hashfile """ & filePath & """ SHA256").StdOut.ReadAll
    Dim lines As Variant: lines = Split(output, vbCrLf)
    If UBound(lines) >= 1 Then
        GetSha256 = Replace(Trim$(CStr(lines(1))), " ", "")
    End If
    Exit Function
Failed:
    GetSha256 = ""
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function
