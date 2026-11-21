Attribute VB_Name = "Mod111_処理ログ"
Option Explicit

Public Sub ファイル処理ログテキストExcel出力()
    Dim message As String
    message = InputBox("ログに記録する内容を入力", "処理ログ")
    If Trim$(message) = "" Then Exit Sub
    Dim mode As String
    mode = InputBox("出力形式を入力: 1=テキスト、2=Excelシート", "処理ログ", "1")
    If mode = "1" Then
        Dim logPath As String
        logPath = ThisWorkbook.Path & "\file_processing.log"
        Dim fileNumber As Integer: fileNumber = FreeFile
        Open logPath For Append As #fileNumber
        Print #fileNumber, Format$(Now, "yyyy/mm/dd hh:mm:ss") & vbTab & message
        Close #fileNumber
        MsgBox "ログへ追記しました。" & vbCrLf & logPath, vbInformation
    ElseIf mode = "2" Then
        Dim ws As Worksheet: Set ws = GetOrCreateSheet("処理ログ")
        Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(r, 1).Value = Now
        ws.Cells(r, 2).Value = message
        ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    End If
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function
