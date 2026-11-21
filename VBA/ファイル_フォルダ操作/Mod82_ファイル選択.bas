Attribute VB_Name = "Mod82_ファイル選択"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 単一ファイル選択を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「単一ファイル選択」を選択して実行する。

' ── 単一ファイル選択 ──
Public Sub 単一ファイル選択()
    Dim selected As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        If .Show = -1 Then selected = .SelectedItems(1)
    End With
    If selected <> "" Then MsgBox selected, vbInformation
End Sub

' ── 複数ファイル選択 ──
Public Sub 複数ファイル選択()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True
        If .Show = -1 Then
            Dim i As Long
            For i = 1 To .SelectedItems.Count
                ActiveSheet.Cells(i, 1).Value = .SelectedItems(i)
            Next i
        End If
    End With
End Sub




