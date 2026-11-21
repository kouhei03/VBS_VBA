Attribute VB_Name = "Mod86_再帰検索"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod86_再帰検索
' - ファイル名: Mod86_再帰検索.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub サブフォルダを含む再帰ファイル検索() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub サブフォルダを含む再帰ファイル検索()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim folderPath As String, keyword As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    keyword = InputBox("検索するファイル名の一部を入力", "再帰検索")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Trim$(keyword) = "" Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet("再帰検索結果")
    ws.Cells.Clear
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long: r = 1
    SearchFiles fso.GetFolder(folderPath), keyword, ws, r
    ws.Columns("A:B").AutoFit
    MsgBox r - 1 & " 件見つかりました。", vbInformation
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub SearchFiles(ByVal folder As Object, ByVal keyword As String, ByVal ws As Worksheet, ByRef r As Long) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub SearchFiles(ByVal folder As Object, ByVal keyword As String, ByVal ws As Worksheet, ByRef r As Long)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim file As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each file In folder.Files
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If InStr(1, file.Name, keyword, vbTextCompare) > 0 Then
            ws.Cells(r, 1).Value = file.Name
            ws.Cells(r, 2).Value = file.Path
            r = r + 1
        End If
    Next file
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim subfolder As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each subfolder In folder.SubFolders
        SearchFiles subfolder, keyword, ws, r
    Next subfolder
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOrCreateSheet Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function

