Attribute VB_Name = "Mod115_テンプレート複製"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作、ブック・シートを処理を実行する。
' ファイルないコード処理詳細解説: テンプレートファイル複製作業ブック生成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「テンプレートファイル複製作業ブック生成」を選択して実行する。

' ── テンプレートファイル複製作業ブック生成 ──
Public Sub テンプレートファイル複製作業ブック生成()
    Dim templatePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Add "Excel", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then Exit Sub
        templatePath = .SelectedItems(1)
    End With
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim name As String
    name = InputBox("作業ブック名を入力", "テンプレート複製", "作業_" & Format$(Now, "yyyymmdd"))
    If Trim$(name) = "" Then Exit Sub
    Dim extension As String: extension = "." & fso.GetExtensionName(templatePath)
    If LCase$(Right$(name, Len(extension))) <> LCase$(extension) Then name = name & extension
    fso.CopyFile templatePath, fso.BuildPath(destinationFolder, name), True
    Workbooks.Open fso.BuildPath(destinationFolder, name)
End Sub




