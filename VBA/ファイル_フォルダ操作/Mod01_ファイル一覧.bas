Attribute VB_Name = "Mod01_ファイル一覧"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod01_ファイル一覧
' - ファイル名: Mod01_ファイル一覧.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' 指定フォルダ内のファイル情報をシートに一覧出力する。
' サブフォルダを含めるかどうかは実行時に選択する。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub フォルダ内ファイル一覧() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub フォルダ内ファイル一覧()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim folderPath As String
    folderPath = SelectFolder("一覧化するフォルダを選択してください")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If folderPath = "" Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim includeSubfolders As Boolean
    includeSubfolders = (MsgBox("サブフォルダ内のファイルも含めますか？", _
                                vbYesNo + vbQuestion) = vbYes)

    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Failed

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fso As Object
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim files As Collection
    Set files = New Collection
    CollectFiles fso.GetFolder(folderPath), files, includeSubfolders

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = GetOrCreateSheet(ThisWorkbook, "ファイル一覧")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "ファイル名"
    ws.Cells(1, 2).Value = "更新日時"
    ws.Cells(1, 3).Value = "サイズ (bytes)"
    ws.Cells(1, 4).Value = "フルパス"

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If files.Count > 0 Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim data() As Variant
        ReDim data(1 To files.Count, 1 To 4)

        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim i As Long
        Dim file As Object
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For i = 1 To files.Count
            Set file = files(i)
            data(i, 1) = file.Name
            data(i, 2) = file.DateLastModified
            data(i, 3) = CDbl(file.Size)
            data(i, 4) = file.Path
        Next i

        ws.Range("A2").Resize(files.Count, 4).Value = data
        ws.Columns(2).NumberFormat = "yyyy/mm/dd hh:mm:ss"
        ws.Columns(3).NumberFormat = "#,##0"
    End If

    ws.Rows(1).Font.Bold = True
    ws.Columns("A:D").AutoFit
    MsgBox files.Count & " 件のファイルを一覧にしました。", vbInformation
    Exit Sub

Failed:
    MsgBox "ファイル一覧の作成に失敗しました: " & Err.Description, vbCritical
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub CollectFiles(ByVal folder As Object, ByRef files As Collection, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub CollectFiles(ByVal folder As Object, ByRef files As Collection, _
                         ByVal includeSubfolders As Boolean)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim file As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each file In folder.Files
        files.Add file
    Next file

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If includeSubfolders Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim subfolder As Object
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For Each subfolder In folder.SubFolders
            CollectFiles subfolder, files, True
        Next subfolder
    End If
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function SelectFolder(ByVal title As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function SelectFolder(ByVal title As String) As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title
        .AllowMultiSelect = False
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show = -1 Then SelectFolder = .SelectedItems(1)
    End With
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOrCreateSheet = wb.Worksheets(sheetName)
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOrCreateSheet Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

