Attribute VB_Name = "Mod64_クロス集計"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod64_クロス集計
' - ファイル名: Mod64_クロス集計.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub クロス集計表を作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub クロス集計表を作成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim rowCol As Long: rowCol = CLng(InputBox("行項目の列番号", "クロス集計", "1"))
    Dim colCol As Long: colCol = CLng(InputBox("列項目の列番号", "クロス集計", "2"))
    Dim valueCol As Long: valueCol = CLng(InputBox("集計値の列番号", "クロス集計", "3"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "クロス集計")
    out.Cells.Clear

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim rowsDic As Object: Set rowsDic = CreateObject("Scripting.Dictionary")
    Dim colsDic As Object: Set colsDic = CreateObject("Scripting.Dictionary")
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(ws)
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim rowKey As String: rowKey = CStr(ws.Cells(r, rowCol).Value)
        Dim colKey As String: colKey = CStr(ws.Cells(r, colCol).Value)
        rowsDic(rowKey) = True: colsDic(colKey) = True
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim key As String: key = rowKey & vbNullChar & colKey
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Not sums.Exists(key) Then sums.Add key, 0
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsNumeric(ws.Cells(r, valueCol).Value) Then sums(key) = sums(key) + CDbl(ws.Cells(r, valueCol).Value)
    Next r
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim c As Long: c = 2
    Dim colName As Variant
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each colName In colsDic.Keys
        out.Cells(1, c).Value = colName
        c = c + 1
    Next colName
    r = 2
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim rowName As Variant
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each rowName In rowsDic.Keys
        out.Cells(r, 1).Value = rowName
        c = 2
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For Each colName In colsDic.Keys
            key = CStr(rowName) & vbNullChar & CStr(colName)
            out.Cells(r, c).Value = IIf(sums.Exists(key), sums(key), 0)
            c = c + 1
        Next colName
        r = r + 1
    Next rowName
End Sub

