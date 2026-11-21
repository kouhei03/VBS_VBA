Attribute VB_Name = "Mod209_ExpectedActual比較"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod209_ExpectedActual比較
' - ファイル名: Mod209_ExpectedActual比較.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ExpectedActualセル単位一括比較() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ExpectedActualセル単位一括比較()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim mismatch As Worksheet: Set mismatch = GetOrCreateSheet
    mismatch.Cells.Clear
    mismatch.Range("A1:D1").Value = Array("セル", "Expected", "Actual", "判定")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long: r = 2
    Dim c As Range
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each c In Selection
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If c.Column = 1 Then
            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim actualCell As Range: Set actualCell = c.Offset(0, 1)
            mismatch.Cells(r, 1).Value = c.Address
            mismatch.Cells(r, 2).Value = c.Value
            mismatch.Cells(r, 3).Value = actualCell.Value
            mismatch.Cells(r, 4).Value = IIf(CStr(c.Value) = CStr(actualCell.Value), "OK", "NG")
            r = r + 1
        End If
    Next c
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function GetOrCreateSheet() As Worksheet の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function GetOrCreateSheet() As Worksheet
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error Resume Next
    ' [ブロック解説: オブジェクト参照の設定]
    ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("比較結果")
    ' [ブロック解説: エラー処理設定]
    ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
    On Error GoTo 0
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If GetOrCreateSheet Is Nothing Then
        ' [ブロック解説: オブジェクト参照の設定]
        ' Workbook、Worksheet、Range、外部COMオブジェクトなどの参照を初期化します。オブジェクトが存在し、必要な参照ライブラリまたはアプリケーションが利用可能か確認します。
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "比較結果"
    End If
End Function

