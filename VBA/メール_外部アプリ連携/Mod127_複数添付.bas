Attribute VB_Name = "Mod127_複数添付"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 複数ファイル添付を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「複数ファイル添付」を選択して実行する。

' ── 複数ファイル添付 ──
Public Sub 複数ファイル添付()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "添付メール")
    mail.Subject = InputBox("件名を入力してください", "添付メール")
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        If .Show = -1 Then
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                mail.Attachments.Add .SelectedItems(i)
            Next i
        End If
    End With
    mail.Save
End Sub




