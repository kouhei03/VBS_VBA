Attribute VB_Name = "Mod00_共通"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod00_共通
' - ファイル名: Mod00_共通.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

'==================================================
' Mod00_共通
' 各モジュールから利用する共通ヘルパー関数
'==================================================

'--------------------------------------------------
' 指定シートの最終行を取得（全列を対象）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function GetLastRow(ByVal ws As Worksheet) As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function GetLastRow(ByVal ws As Worksheet) As Long
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim rng As Range
    Set rng = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If rng Is Nothing Then
        GetLastRow = 0
    Else
        GetLastRow = rng.Row
    End If
End Function

'--------------------------------------------------
' 指定シートの最終列を取得（全行を対象）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function GetLastCol(ByVal ws As Worksheet) As Long の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function GetLastCol(ByVal ws As Worksheet) As Long
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim rng As Range
    Set rng = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If rng Is Nothing Then
        GetLastCol = 0
    Else
        GetLastCol = rng.Column
    End If
End Function

'--------------------------------------------------
' シートの存在確認
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set ws = wb.Worksheets(sheetName)
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'--------------------------------------------------
' シートを取得（なければ新規作成）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If SheetExists(wb, sheetName) Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    Else
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

'--------------------------------------------------
' シート名として使えない文字を置換
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function SafeSheetName(ByVal s As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function SafeSheetName(ByVal s As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim badChars As Variant
    Dim i As Long
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "_")
    Next i
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Len(s) > 31 Then s = Left$(s, 31)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Len(s) = 0 Then s = "_blank_"
    SafeSheetName = s
End Function

'--------------------------------------------------
' フォルダ選択ダイアログ
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function SelectFolder(Optional ByVal title As String = "フォルダを選択してください") As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function SelectFolder(Optional ByVal title As String = "フォルダを選択してください") As String
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        .AllowMultiSelect = False
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""   ' キャンセル時は空文字
        End If
    End With
End Function

'--------------------------------------------------
' 高速化設定の ON/OFF
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub SpeedUp(ByVal enable As Boolean) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub SpeedUp(ByVal enable As Boolean)
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Application
        .ScreenUpdating = Not enable
        .EnableEvents = Not enable
        .Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not enable
    End With
End Sub

