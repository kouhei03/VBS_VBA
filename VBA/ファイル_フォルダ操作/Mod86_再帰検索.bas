Attribute VB_Name = "Mod86_再帰検索"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: サブフォルダを含む再帰ファイル検索を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「サブフォルダを含む再帰ファイル検索」を選択して実行する。

' ── サブフォルダを含む再帰ファイル検索 ──
Public Sub サブフォルダを含む再帰ファイル検索()
    Dim folderPath As String, keyword As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    keyword = InputBox("検索するファイル名の一部を入力", "再帰検索")
    If Trim$(keyword) = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("再帰検索結果")
    ws.Cells.Clear
    Dim r As Long: r = 1
    SearchFiles fso.GetFolder(folderPath), keyword, ws, r
    ws.Columns("A:B").AutoFit
    MsgBox r - 1 & " 件見つかりました。", vbInformation
End Sub

' ── SearchFiles ──
Private Sub SearchFiles(ByVal folder As Object, ByVal keyword As String, ByVal ws As Worksheet, ByRef r As Long)
    Dim file As Object
    For Each file In folder.Files
        If InStr(1, file.Name, keyword, vbTextCompare) > 0 Then
            ws.Cells(r, 1).Value = file.Name
            ws.Cells(r, 2).Value = file.Path
            r = r + 1
        End If
    Next file
    Dim subfolder As Object
    For Each subfolder In folder.SubFolders
        SearchFiles subfolder, keyword, ws, r
    Next subfolder
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




