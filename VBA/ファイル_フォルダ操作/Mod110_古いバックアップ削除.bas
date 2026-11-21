Attribute VB_Name = "Mod110_古いバックアップ削除"
Option Explicit
' ファイル処理概要: ファイル・フォルダを操作を実行する。
' ファイルないコード処理詳細解説: 古いバックアップファイル自動削除を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、Windowsファイルシステム、必要な実行権限。
' ファイル実行例: ExcelのAlt+F8から「古いバックアップファイル自動削除」を選択して実行する。

' ── 古いバックアップファイル自動削除 ──
Public Sub 古いバックアップファイル自動削除()
    Dim backupFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "バックアップフォルダを選択"
        If .Show <> -1 Then Exit Sub
        backupFolder = .SelectedItems(1)
    End With
    Dim keepDays As Variant
    keepDays = Application.InputBox("保持日数を入力してください", "世代管理", 30, , , , , 1)
    If keepDays = False Or Not IsNumeric(keepDays) Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim deletedCount As Long: deletedCount = 0
    Dim cutoff As Date: cutoff = DateAdd("d", -CLng(keepDays), Now)
    Dim file As Object
    For Each file In fso.GetFolder(backupFolder).Files
        If file.DateLastModified < cutoff Then
            file.Delete True
            deletedCount = deletedCount + 1
        End If
    Next file
    MsgBox deletedCount & " 件削除しました。", vbInformation
End Sub




