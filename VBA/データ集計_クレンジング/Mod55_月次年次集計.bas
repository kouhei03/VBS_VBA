Attribute VB_Name = "Mod55_月次年次集計"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod55_月次年次集計
' - ファイル名: Mod55_月次年次集計.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 月次集計表を作成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 月次集計表を作成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dateCol As Long: dateCol = CLng(InputBox("日付列番号", "月次集計", "1"))
    Dim valueCol As Long: valueCol = CLng(InputBox("集計列番号", "月次集計", "3"))
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "月次集計")
    out.Cells.Clear
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(ws)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsDate(ws.Cells(r, dateCol).Value) Then
            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim key As String: key = Format$(CDate(ws.Cells(r, dateCol).Value), "yyyy/mm")
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If Not dic.Exists(key) Then dic.Add key, 0
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If IsNumeric(ws.Cells(r, valueCol).Value) Then dic(key) = dic(key) + CDbl(ws.Cells(r, valueCol).Value)
        End If
    Next r
    out.Cells(1, 1).Value = "年月": out.Cells(1, 2).Value = "合計"
    r = 2
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim k As Variant
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each k In dic.Keys
        out.Cells(r, 1).Value = k
        out.Cells(r, 2).Value = dic(k)
        r = r + 1
    Next k
End Sub

