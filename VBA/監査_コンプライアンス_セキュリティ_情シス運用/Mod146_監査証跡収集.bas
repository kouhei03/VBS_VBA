Attribute VB_Name = "Mod146_監査証跡収集"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 内部外部監査証跡ファイル収集索引一覧を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「内部外部監査証跡ファイル収集索引一覧」を選択して実行する。

' ── 内部外部監査証跡ファイル収集索引一覧 ──
Public Sub 内部外部監査証跡ファイル収集索引一覧()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("ファイル名", "更新日時", "サイズ", "パス")
    Dim r As Long: r = 2
    Dim file As Object
    For Each file In fso.GetFolder(folderPath).Files
        output.Cells(r, 1).Value = file.Name
        output.Cells(r, 2).Value = file.DateLastModified
        output.Cells(r, 3).Value = file.Size
        output.Cells(r, 4).Value = file.Path
        r = r + 1
    Next file
    output.Columns("A:D").AutoFit
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("監査証跡索引")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "監査証跡索引"
    End If
End Function




