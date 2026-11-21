Attribute VB_Name = "Mod02_一括リネーム"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod02_一括リネーム
' - ファイル名: Mod02_一括リネーム.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' 選択フォルダ直下のファイル名に接頭辞と置換ルールを適用する。
' 実行前に変更一覧を表示し、確認後に一時名を経由してリネームする。
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ファイル名一括変更() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ファイル名一括変更()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim folderPath As String
    folderPath = SelectFolder("リネームするファイルがあるフォルダを選択してください")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If folderPath = "" Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim prefix As String
    Dim canceled As Boolean
    prefix = PromptText("ファイル名の先頭に付ける文字を入力してください。" & vbCrLf & _
                        "不要な場合は空欄のままにします。", "接頭辞", canceled)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If canceled Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim findText As String
    findText = PromptText("ファイル名から置換する文字を入力してください。" & vbCrLf & _
                          "置換しない場合は空欄にします。", "置換対象", canceled)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If canceled Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim replaceText As String
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If findText <> "" Then
        replaceText = PromptText("置換後の文字を入力してください。", "置換後", canceled)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If canceled Then Exit Sub
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim addNumber As Boolean
    addNumber = (MsgBox("先頭に連番（001_、002_...）を付けますか？", _
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
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim sourceFiles As Collection
    Set sourceFiles = New Collection
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim file As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each file In folder.Files
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Left$(file.Name, 2) <> "~$" And _
           StrComp(file.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            sourceFiles.Add file
        End If
    Next file

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If sourceFiles.Count = 0 Then
        MsgBox "対象ファイルがありません。", vbInformation
        Exit Sub
    End If

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim oldPaths() As String, newPaths() As String, tempPaths() As String
    Dim staged() As Boolean, finalized() As Boolean
    Dim errorText As String
    ReDim oldPaths(1 To sourceFiles.Count)
    ReDim newPaths(1 To sourceFiles.Count)
    ReDim tempPaths(1 To sourceFiles.Count)
    ReDim staged(1 To sourceFiles.Count)
    ReDim finalized(1 To sourceFiles.Count)

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    Dim oldName As String, newName As String, baseName As String, extension As String
    Dim targetNames As Object, sourceNames As Object
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set targetNames = CreateObject("Scripting.Dictionary")
    targetNames.CompareMode = vbTextCompare
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set sourceNames = CreateObject("Scripting.Dictionary")
    sourceNames.CompareMode = vbTextCompare

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        Set file = sourceFiles(i)
        oldPaths(i) = file.Path
        sourceNames(file.Path) = True
    Next i

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim numberWidth As Long
    numberWidth = Len(CStr(sourceFiles.Count))
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If numberWidth < 3 Then numberWidth = 3

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        Set file = sourceFiles(i)
        oldName = file.Name
        baseName = fso.GetBaseName(oldName)
        extension = fso.GetExtensionName(oldName)

        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If findText <> "" Then baseName = Replace(baseName, findText, replaceText, 1, -1, vbTextCompare)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If addNumber Then baseName = Format$(i, String$(numberWidth, "0")) & "_" & baseName
        newName = prefix & baseName
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If extension <> "" Then newName = newName & "." & extension

        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Not IsValidFileName(newName) Then
            Err.Raise vbObjectError + 1000, , "使用できない文字を含むファイル名です: " & newName
        End If

        newPaths(i) = fso.BuildPath(folderPath, newName)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If targetNames.Exists(newPaths(i)) Then
            Err.Raise vbObjectError + 1001, , "変更後のファイル名が重複しています: " & newName
        End If
        targetNames.Add newPaths(i), True
    Next i

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If fso.FileExists(newPaths(i)) And Not sourceNames.Exists(newPaths(i)) Then
            Err.Raise vbObjectError + 1002, , "同名のファイルが既にあります: " & fso.GetFileName(newPaths(i))
        End If
        tempPaths(i) = fso.BuildPath(folderPath, "~vba_rename_" & Format$(i, "000000") & "_" & _
                                     Format$(Now, "yymmddhhnnss") & ".tmp")
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If fso.FileExists(tempPaths(i)) Then
            Err.Raise vbObjectError + 1003, , "一時ファイル名が既に存在します。再実行してください。"
        End If
    Next i

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = GetOrCreateSheet(ThisWorkbook, "リネーム確認")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "変更前"
    ws.Cells(1, 2).Value = "変更後"
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        ws.Cells(i + 1, 1).Value = fso.GetFileName(oldPaths(i))
        ws.Cells(i + 1, 2).Value = fso.GetFileName(newPaths(i))
    Next i
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:B").AutoFit
    ws.Activate

    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If MsgBox(sourceFiles.Count & " 件のファイル名を変更します。" & vbCrLf & _
              "「リネーム確認」シートの一覧を確認してから実行してください。", _
              vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        fso.MoveFile oldPaths(i), tempPaths(i)
        staged(i) = True
    Next i

    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        fso.MoveFile tempPaths(i), newPaths(i)
        finalized(i) = True
        staged(i) = False
    Next i

    MsgBox sourceFiles.Count & " 件のファイル名を変更しました。", vbInformation
    Exit Sub

Failed:
    errorText = Err.Description
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If finalized(i) Then
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If fso.FileExists(newPaths(i)) Then fso.MoveFile newPaths(i), tempPaths(i)
            staged(i) = True
        End If
    Next i
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = 1 To sourceFiles.Count
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If staged(i) And fso.FileExists(tempPaths(i)) Then
            fso.MoveFile tempPaths(i), oldPaths(i)
        End If
    Next i
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    MsgBox "リネームに失敗しました: " & errorText & vbCrLf & _
           "可能な範囲で元のファイル名に戻しました。フォルダ内をご確認ください。", vbCritical
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function PromptText(ByVal prompt As String, ByVal title As String, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function PromptText(ByVal prompt As String, ByVal title As String, _
                            ByRef canceled As Boolean) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim answer As Variant
    answer = Application.InputBox(prompt, title, Type:=2)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If VarType(answer) = vbBoolean Then
        canceled = True
    Else
        canceled = False
        PromptText = CStr(answer)
    End If
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function IsValidFileName(ByVal fileName As String) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function IsValidFileName(ByVal fileName As String) As Boolean
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Len(fileName) = 0 Or fileName = "." Or fileName = ".." Then Exit Function
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Right$(fileName, 1) = "." Or Right$(fileName, 1) = " " Then Exit Function

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim invalidChars As Variant
    invalidChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim i As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(invalidChars) To UBound(invalidChars)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If InStr(fileName, invalidChars(i)) > 0 Then Exit Function
    Next i
    IsValidFileName = True
End Function

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

