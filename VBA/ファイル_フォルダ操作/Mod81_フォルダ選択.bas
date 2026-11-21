Attribute VB_Name = "Mod81_フォルダ選択"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: フォルダ選択パス取得を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「フォルダ選択パス取得」を選択して実行する。

' ── フォルダ選択パス取得 ──
Public Sub フォルダ選択パス取得()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択してください"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
            MsgBox "選択フォルダ: " & vbCrLf & folderPath, vbInformation
        End If
    End With
End Sub




