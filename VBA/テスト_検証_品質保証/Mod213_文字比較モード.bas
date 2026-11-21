Attribute VB_Name = "Mod213_文字比較モード"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod213_文字比較モード
' - ファイル名: Mod213_文字比較モード.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 空白改行大小文字比較モード切替() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 空白改行大小文字比較モード切替()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim mode As String
    mode = InputBox("モード: 1=厳密、2=空白改行無視、3=大小文字無視", "文字比較")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Selection.Columns.Count < 2 Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim c As Range
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each c In Selection.Columns(1).Cells
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If c.Row > 0 And c.Row <= Selection.Rows.Count + Selection.Row - 1 Then
            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim rightValue As String: rightValue = CStr(c.Offset(0, 1).Value)
            Dim leftValue As String: leftValue = CStr(c.Value)
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If mode = "2" Then
                leftValue = Replace(Replace(leftValue, " ", ""), vbCrLf, "")
                rightValue = Replace(Replace(rightValue, " ", ""), vbCrLf, "")
            ElseIf mode = "3" Then
                leftValue = LCase$(leftValue)
                rightValue = LCase$(rightValue)
            End If
            c.Offset(0, 2).Value = IIf(leftValue = rightValue, "OK", "NG")
        End If
    Next c
End Sub

