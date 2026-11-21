Attribute VB_Name = "Mod66_欠損値リスト"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod66_欠損値リスト
' - ファイル名: Mod66_欠損値リスト.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 欠損値をリストアップ() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 欠損値をリストアップ()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "欠損値リスト")
    out.Cells.Clear
    out.Cells(1, 1).Value = "行": out.Cells(1, 2).Value = "列": out.Cells(1, 3).Value = "列名"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outputRow As Long: outputRow = 2
    Dim r As Long, c As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(ws)
        ' [ブロック解説: 繰り返し処理]
        ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
        For c = 1 To GetLastCol(ws)
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) = 0 Then
                out.Cells(outputRow, 1).Value = r
                out.Cells(outputRow, 2).Value = c
                out.Cells(outputRow, 3).Value = ws.Cells(1, c).Value
                outputRow = outputRow + 1
            End If
        Next c
    Next r
End Sub

