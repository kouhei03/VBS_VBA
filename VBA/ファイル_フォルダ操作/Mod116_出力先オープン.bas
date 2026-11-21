Attribute VB_Name = "Mod116_出力先オープン"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 出力先フォルダエクスプローラ表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「出力先フォルダエクスプローラ表示」を選択して実行する。

' ── 出力先フォルダエクスプローラ表示 ──
Public Sub 出力先フォルダエクスプローラ表示()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "表示する出力先フォルダを選択"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Shell "explorer.exe """ & folderPath & """", vbNormalFocus
End Sub




