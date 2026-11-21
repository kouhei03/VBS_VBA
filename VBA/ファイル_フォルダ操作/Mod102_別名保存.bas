Attribute VB_Name = "Mod102_別名保存"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: ブック別名保存日付バージョン付与を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「ブック別名保存日付バージョン付与」を選択して実行する。

' ── ブック別名保存日付バージョン付与 ──
Public Sub ブック別名保存日付バージョン付与()
    Dim mode As String
    mode = InputBox("付与形式を入力: 1=日付、2=バージョン、3=両方", "別名保存", "1")
    Dim suffix As String
    If mode = "1" Then suffix = "_" & Format$(Now, "yyyymmdd")
    If mode = "2" Then suffix = "_v" & Format$(1, "00")
    If mode = "3" Then suffix = "_v1_" & Format$(Now, "yyyymmdd")
    If suffix = "" Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseName As String: baseName = fso.GetBaseName(ThisWorkbook.Name)
    Dim extension As String: extension = fso.GetExtensionName(ThisWorkbook.Name)
    Dim destinationFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destinationFolder = .SelectedItems(1)
    End With
    Dim destination As String
    destination = fso.BuildPath(destinationFolder, baseName & suffix & "." & extension)
    ThisWorkbook.SaveCopyAs destination
    MsgBox "別名保存しました。" & vbCrLf & destination, vbInformation
End Sub




