Attribute VB_Name = "Mod109_バックアップ世代"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: バックアップフォルダへ世代管理保存を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「バックアップフォルダへ世代管理保存」を選択して実行する。

' ── バックアップフォルダへ世代管理保存 ──
Public Sub バックアップフォルダへ世代管理保存()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(ThisWorkbook.FullName) Then
        MsgBox "保存済みブックに対して実行してください。", vbExclamation
        Exit Sub
    End If
    Dim backupParent As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "バックアップ先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        backupParent = .SelectedItems(1)
    End With
    Dim generationFolder As String
    generationFolder = fso.BuildPath(backupParent, Format$(Now, "yyyymmdd_hhnnss"))
    fso.CreateFolder generationFolder
    fso.CopyFile ThisWorkbook.FullName, fso.BuildPath(generationFolder, fso.GetFileName(ThisWorkbook.FullName)), True
    MsgBox "世代バックアップを保存しました。" & vbCrLf & generationFolder, vbInformation
End Sub




