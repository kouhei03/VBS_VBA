Attribute VB_Name = "Mod02_一括リネーム"
Option Explicit

' 選択フォルダ直下のファイル名に接頭辞と置換ルールを適用する。
' 実行前に変更一覧を表示し、確認後に一時名を経由してリネームする。
Public Sub ファイル名一括変更()
    Dim folderPath As String
    folderPath = SelectFolder("リネームするファイルがあるフォルダを選択してください")
    If folderPath = "" Then Exit Sub

    Dim prefix As String
    Dim canceled As Boolean
    prefix = PromptText("ファイル名の先頭に付ける文字を入力してください。" & vbCrLf & _
                        "不要な場合は空欄のままにします。", "接頭辞", canceled)
    If canceled Then Exit Sub

    Dim findText As String
    findText = PromptText("ファイル名から置換する文字を入力してください。" & vbCrLf & _
                          "置換しない場合は空欄にします。", "置換対象", canceled)
    If canceled Then Exit Sub

    Dim replaceText As String
    If findText <> "" Then
        replaceText = PromptText("置換後の文字を入力してください。", "置換後", canceled)
        If canceled Then Exit Sub
    End If

    Dim addNumber As Boolean
    addNumber = (MsgBox("先頭に連番（001_、002_...）を付けますか？", _
                        vbYesNo + vbQuestion) = vbYes)

    On Error GoTo Failed

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    Dim sourceFiles As Collection
    Set sourceFiles = New Collection
    Dim file As Object
    For Each file In folder.Files
        If Left$(file.Name, 2) <> "~$" And _
           StrComp(file.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            sourceFiles.Add file
        End If
    Next file

    If sourceFiles.Count = 0 Then
        MsgBox "対象ファイルがありません。", vbInformation
        Exit Sub
    End If

    Dim oldPaths() As String, newPaths() As String, tempPaths() As String
    Dim staged() As Boolean, finalized() As Boolean
    Dim errorText As String
    ReDim oldPaths(1 To sourceFiles.Count)
    ReDim newPaths(1 To sourceFiles.Count)
    ReDim tempPaths(1 To sourceFiles.Count)
    ReDim staged(1 To sourceFiles.Count)
    ReDim finalized(1 To sourceFiles.Count)

    Dim i As Long
    Dim oldName As String, newName As String, baseName As String, extension As String
    Dim targetNames As Object, sourceNames As Object
    Set targetNames = CreateObject("Scripting.Dictionary")
    targetNames.CompareMode = vbTextCompare
    Set sourceNames = CreateObject("Scripting.Dictionary")
    sourceNames.CompareMode = vbTextCompare

    For i = 1 To sourceFiles.Count
        Set file = sourceFiles(i)
        oldPaths(i) = file.Path
        sourceNames(file.Path) = True
    Next i

    Dim numberWidth As Long
    numberWidth = Len(CStr(sourceFiles.Count))
    If numberWidth < 3 Then numberWidth = 3

    For i = 1 To sourceFiles.Count
        Set file = sourceFiles(i)
        oldName = file.Name
        baseName = fso.GetBaseName(oldName)
        extension = fso.GetExtensionName(oldName)

        If findText <> "" Then baseName = Replace(baseName, findText, replaceText, 1, -1, vbTextCompare)
        If addNumber Then baseName = Format$(i, String$(numberWidth, "0")) & "_" & baseName
        newName = prefix & baseName
        If extension <> "" Then newName = newName & "." & extension

        If Not IsValidFileName(newName) Then
            Err.Raise vbObjectError + 1000, , "使用できない文字を含むファイル名です: " & newName
        End If

        newPaths(i) = fso.BuildPath(folderPath, newName)
        If targetNames.Exists(newPaths(i)) Then
            Err.Raise vbObjectError + 1001, , "変更後のファイル名が重複しています: " & newName
        End If
        targetNames.Add newPaths(i), True
    Next i

    For i = 1 To sourceFiles.Count
        If fso.FileExists(newPaths(i)) And Not sourceNames.Exists(newPaths(i)) Then
            Err.Raise vbObjectError + 1002, , "同名のファイルが既にあります: " & fso.GetFileName(newPaths(i))
        End If
        tempPaths(i) = fso.BuildPath(folderPath, "~vba_rename_" & Format$(i, "000000") & "_" & _
                                     Format$(Now, "yymmddhhnnss") & ".tmp")
        If fso.FileExists(tempPaths(i)) Then
            Err.Raise vbObjectError + 1003, , "一時ファイル名が既に存在します。再実行してください。"
        End If
    Next i

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "リネーム確認")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "変更前"
    ws.Cells(1, 2).Value = "変更後"
    For i = 1 To sourceFiles.Count
        ws.Cells(i + 1, 1).Value = fso.GetFileName(oldPaths(i))
        ws.Cells(i + 1, 2).Value = fso.GetFileName(newPaths(i))
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:B").AutoFit
    ws.Activate

    If MsgBox(sourceFiles.Count & " 件のファイル名を変更します。" & vbCrLf & _
              "「リネーム確認」シートの一覧を確認してから実行してください。", _
              vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    For i = 1 To sourceFiles.Count
        fso.MoveFile oldPaths(i), tempPaths(i)
        staged(i) = True
    Next i

    For i = 1 To sourceFiles.Count
        fso.MoveFile tempPaths(i), newPaths(i)
        finalized(i) = True
        staged(i) = False
    Next i

    MsgBox sourceFiles.Count & " 件のファイル名を変更しました。", vbInformation
    Exit Sub

Failed:
    errorText = Err.Description
    On Error Resume Next
    For i = 1 To sourceFiles.Count
        If finalized(i) Then
            If fso.FileExists(newPaths(i)) Then fso.MoveFile newPaths(i), tempPaths(i)
            staged(i) = True
        End If
    Next i
    For i = 1 To sourceFiles.Count
        If staged(i) And fso.FileExists(tempPaths(i)) Then
            fso.MoveFile tempPaths(i), oldPaths(i)
        End If
    Next i
    On Error GoTo 0
    MsgBox "リネームに失敗しました: " & errorText & vbCrLf & _
           "可能な範囲で元のファイル名に戻しました。フォルダ内をご確認ください。", vbCritical
End Sub

Private Function PromptText(ByVal prompt As String, ByVal title As String, _
                            ByRef canceled As Boolean) As String
    Dim answer As Variant
    answer = Application.InputBox(prompt, title, Type:=2)
    If VarType(answer) = vbBoolean Then
        canceled = True
    Else
        canceled = False
        PromptText = CStr(answer)
    End If
End Function

Private Function IsValidFileName(ByVal fileName As String) As Boolean
    If Len(fileName) = 0 Or fileName = "." Or fileName = ".." Then Exit Function
    If Right$(fileName, 1) = "." Or Right$(fileName, 1) = " " Then Exit Function

    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    Dim i As Long
    For i = LBound(invalidChars) To UBound(invalidChars)
        If InStr(fileName, invalidChars(i)) > 0 Then Exit Function
    Next i
    IsValidFileName = True
End Function

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
