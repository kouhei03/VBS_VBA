Attribute VB_Name = "Mod113_空フォルダ"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 空フォルダ検出削除を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「空フォルダ検出削除」を選択して実行する。

' ── 空フォルダ検出削除 ──
Public Sub 空フォルダ検出削除()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim deleteEmpty As Boolean
    deleteEmpty = MsgBox("空フォルダを検出し削除しますか？", vbYesNo + vbQuestion) = vbYes
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("空フォルダ")
    ws.Cells.Clear
    Dim r As Long: r = 1
    FindEmptyFolders fso.GetFolder(folderPath), ws, r, deleteEmpty
    ws.Columns("A:A").AutoFit
End Sub

' ── FindEmptyFolders ──
Private Sub FindEmptyFolders(ByVal folder As Object, ByVal ws As Worksheet, ByRef r As Long, ByVal deleteEmpty As Boolean)
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        FindEmptyFolders subfolder, ws, r, deleteEmpty
    Next subfolder
    If folder.Files.Count = 0 And folder.SubFolders.Count = 0 Then
        ws.Cells(r, 1).Value = folder.Path
        r = r + 1
        If deleteEmpty Then folder.Delete True
    End If
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function




