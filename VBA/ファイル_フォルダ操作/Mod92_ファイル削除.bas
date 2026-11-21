Attribute VB_Name = "Mod92_ファイル削除"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ファイル削除ゴミ箱完全削除を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ファイル削除ゴミ箱完全削除」を選択して実行する。

' ── ファイル削除ゴミ箱完全削除 ──
Public Sub ファイル削除ゴミ箱完全削除()
    Dim filePath As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Dim mode As String
    mode = InputBox("削除方式を入力: 1=ゴミ箱、2=完全削除", "ファイル削除", "1")
    If mode = "1" Then
        If MsgBox("ゴミ箱へ移動しますか？" & vbCrLf & filePath, vbYesNo + vbQuestion) = vbYes Then
            With CreateObject("Shell.Application")
                .Namespace(0).ParseName(filePath).InvokeVerb("delete")
            End With
        End If
    ElseIf mode = "2" Then
        If MsgBox("完全削除します。元に戻せません。" & vbCrLf & filePath, vbYesNo + vbCritical) = vbYes Then
            Kill filePath
        End If
    End If
End Sub




