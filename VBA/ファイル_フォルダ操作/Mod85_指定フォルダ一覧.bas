Attribute VB_Name = "Mod85_指定フォルダ一覧"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 指定フォルダ内ファイル一覧取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「指定フォルダ内ファイル一覧取得」を選択して実行する。

' ── 指定フォルダ内ファイル一覧取得 ──
Public Sub 指定フォルダ内ファイル一覧取得()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(folderPath)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("指定フォルダ一覧")
    ws.Cells.Clear

    Dim r As Long: r = 1
    Dim file As Object
    For Each file In folder.Files
        ws.Cells(r, 1).Value = file.Name
        ws.Cells(r, 2).Value = file.Path
        r = r + 1
    Next file
    ws.Columns("A:B").AutoFit
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




