Attribute VB_Name = "Mod01_ファイル統合"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod01_ファイル統合
' - ファイル名: Mod01_ファイル統合.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

'==================================================
' Mod01_ファイル統合
' フォルダ内の複数Excelファイル / 複数シートを
' 1つのマスターシートに自動集約する
'
' 前提:
'   - 各ファイル・各シートは1行目がヘッダーの表形式
'   - ヘッダー構成は全ファイルで同一
'==================================================

' [ブロック解説: 変数・定数・設定値の宣言]
' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
Private Const MASTER_SHEET As String = "マスター"

'--------------------------------------------------
' メイン: フォルダ内の全Excelファイルを統合
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 複数ファイル統合() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 複数ファイル統合()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim folderPath As String
    folderPath = SelectFolder("統合するExcelファイルが入ったフォルダを選択")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If folderPath = "" Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim wsMaster As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set wsMaster = GetOrCreateSheet(ThisWorkbook, MASTER_SHEET)
    wsMaster.Cells.Clear

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim fileName As String
    Dim wb As Workbook
    Dim fileCount As Long
    Dim rowCount As Long

    fileName = Dir(folderPath & Application.PathSeparator & "*.xls*")
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While fileName <> ""
        ' 自分自身・一時ファイルはスキップ
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If fileName <> ThisWorkbook.Name And Left$(fileName, 2) <> "~$" Then
            ' [ブロック解説: オブジェクト参照の設定]
            ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
            Set wb = Workbooks.Open( _
                fileName:=folderPath & Application.PathSeparator & fileName, _
                ReadOnly:=True, UpdateLinks:=0)

            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim ws As Worksheet
            ' [ブロック解説: 繰り返し処理]
            ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
            For Each ws In wb.Worksheets
                rowCount = rowCount + CopyToMaster(ws, wsMaster, fileName)
            Next ws

            wb.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If
        fileName = Dir()
    Loop

    wsMaster.Columns.AutoFit

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox fileCount & " 個のファイルから " & rowCount & " 行を統合しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' メイン: このブック内の全シートを統合
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 複数シート統合() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 複数シート統合()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim wsMaster As Worksheet
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set wsMaster = GetOrCreateSheet(ThisWorkbook, MASTER_SHEET)
    wsMaster.Cells.Clear

    SpeedUp True
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo Finally

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet
    Dim rowCount As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each ws In ThisWorkbook.Worksheets
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If ws.Name <> MASTER_SHEET Then
            rowCount = rowCount + CopyToMaster(ws, wsMaster, ws.Name)
        End If
    Next ws

    wsMaster.Columns.AutoFit

Finally:
    SpeedUp False
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox rowCount & " 行をマスターシートに統合しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' 1シート分のデータをマスターへコピー
' 戻り値: コピーした行数（ヘッダー除く）
'--------------------------------------------------
' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function CopyToMaster(ByVal wsSrc As Worksheet, _ の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function CopyToMaster(ByVal wsSrc As Worksheet, _
                              ByVal wsMaster As Worksheet, _
                              ByVal sourceName As String) As Long
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc)
    lastCol = GetLastCol(wsSrc)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If lastRow < 2 Then Exit Function   ' ヘッダーのみ / 空シート

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim destRow As Long
    destRow = GetLastRow(wsMaster)

    ' マスターが空ならヘッダーをコピーし「取込元」列を追加
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If destRow = 0 Then
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy _
            wsMaster.Cells(1, 1)
        wsMaster.Cells(1, lastCol + 1).Value = "取込元"
        destRow = 1
    End If

    ' データ部をコピー
    wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy
    wsMaster.Cells(destRow + 1, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' 取込元を記録
    wsMaster.Range(wsMaster.Cells(destRow + 1, lastCol + 1), _
                   wsMaster.Cells(destRow + lastRow - 1, lastCol + 1)).Value = sourceName

    CopyToMaster = lastRow - 1
End Function

