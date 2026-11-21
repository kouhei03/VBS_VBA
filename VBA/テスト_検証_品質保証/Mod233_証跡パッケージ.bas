Attribute VB_Name = "Mod233_証跡パッケージ"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: テストエビデンス証跡パッケージ化を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テストエビデンス証跡パッケージ化」を選択して実行する。

' ── テストエビデンス証跡パッケージ化 ──
Public Sub テストエビデンス証跡パッケージ化()
    Dim folderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("ブック名", "シート名", "保存日時")
    output.Cells(2, 1).Value = ThisWorkbook.Name
    output.Cells(2, 2).Value = ActiveSheet.Name
    output.Cells(2, 3).Value = Now
    Dim package As Workbook: Set package = Workbooks.Add(1)
    output.Copy Before:=package.Worksheets(1)
    package.Worksheets(1).Delete
    package.SaveAs folderPath & "\証跡_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx", xlOpenXMLWorkbook
    package.Close False
    MsgBox "テスト証跡パッケージを保存しました。", vbInformation
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("証跡パッケージ")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "証跡パッケージ"
    End If
End Function




