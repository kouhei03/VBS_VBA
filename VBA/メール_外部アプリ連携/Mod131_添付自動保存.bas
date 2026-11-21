Attribute VB_Name = "Mod131_添付自動保存"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 受信メール添付ファイル自動保存を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「受信メール添付ファイル自動保存」を選択して実行する。

' ── 受信メール添付ファイル自動保存 ──
Public Sub 受信メール添付ファイル自動保存()
    Dim destination As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        destination = .SelectedItems(1)
    End With
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim inbox As Object: Set inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    Dim count As Long
    Dim item As Object
    For Each item In inbox.Items
        Dim attachment As Object
        For Each attachment In item.Attachments
            attachment.SaveAsFile destination & "\" & attachment.FileName
            count = count + 1
        Next attachment
    Next item
    MsgBox count & " 件保存しました。", vbInformation
End Sub




