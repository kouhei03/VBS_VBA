Attribute VB_Name = "Mod180_アドイン配置"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: アドインインストール有効化自動配置補助を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「アドインインストール有効化自動配置補助」を選択して実行する。

' ── アドインインストール有効化自動配置補助 ──
Public Sub アドインインストール有効化自動配置補助()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Add "Excel Add-in", "*.xlam"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    AddIns.Add fileName:=filePath, CopyFile:=True
    AddIns(AddIns.Count).Installed = True
    MsgBox "アドインを追加しました。", vbInformation
End Sub




