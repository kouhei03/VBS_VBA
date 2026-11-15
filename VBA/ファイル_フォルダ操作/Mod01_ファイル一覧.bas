Attribute VB_Name = "Mod01_ファイル一覧"
Option Explicit

' 指定フォルダ内のファイル情報をシートに一覧出力する。
' サブフォルダを含めるかどうかは実行時に選択する。
Public Sub フォルダ内ファイル一覧()
    Dim folderPath As String
    folderPath = SelectFolder("一覧化するフォルダを選択してください")
    If folderPath = "" Then Exit Sub

    Dim includeSubfolders As Boolean
    includeSubfolders = (MsgBox("サブフォルダ内のファイルも含めますか？", _
                                vbYesNo + vbQuestion) = vbYes)

    On Error GoTo Failed

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim files As Collection
    Set files = New Collection
    CollectFiles fso.GetFolder(folderPath), files, includeSubfolders

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "ファイル一覧")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "ファイル名"
    ws.Cells(1, 2).Value = "更新日時"
    ws.Cells(1, 3).Value = "サイズ (bytes)"
    ws.Cells(1, 4).Value = "フルパス"

    If files.Count > 0 Then
        Dim data() As Variant
        ReDim data(1 To files.Count, 1 To 4)

        Dim i As Long
        Dim file As Object
        For i = 1 To files.Count
            Set file = files(i)
            data(i, 1) = file.Name
            data(i, 2) = file.DateLastModified
            data(i, 3) = CDbl(file.Size)
            data(i, 4) = file.Path
        Next i

        ws.Range("A2").Resize(files.Count, 4).Value = data
        ws.Columns(2).NumberFormat = "yyyy/mm/dd hh:mm:ss"
        ws.Columns(3).NumberFormat = "#,##0"
    End If

    ws.Rows(1).Font.Bold = True
    ws.Columns("A:D").AutoFit
    MsgBox files.Count & " 件のファイルを一覧にしました。", vbInformation
    Exit Sub

Failed:
    MsgBox "ファイル一覧の作成に失敗しました: " & Err.Description, vbCritical
End Sub

Private Sub CollectFiles(ByVal folder As Object, ByRef files As Collection, _
                         ByVal includeSubfolders As Boolean)
    Dim file As Object
    For Each file In folder.Files
        files.Add file
    Next file

    If includeSubfolders Then
        Dim subfolder As Object
        For Each subfolder In folder.SubFolders
            CollectFiles subfolder, files, True
        Next subfolder
    End If
End Sub

Private Function SelectFolder(ByVal title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title
        .AllowMultiSelect = False
        If .Show = -1 Then SelectFolder = .SelectedItems(1)
    End With
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function
