Attribute VB_Name = "Mod73_累計構成比前月比"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod73_累計構成比前月比
' - ファイル名: Mod73_累計構成比前月比.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 累計構成比前月比を算出() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 累計構成比前月比を算出()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim valueCol As Long: valueCol = CLng(InputBox("数値列番号", "派生指標", "3"))
    Dim total As Double
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(ws)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsNumeric(ws.Cells(r, valueCol).Value) Then total = total + CDbl(ws.Cells(r, valueCol).Value)
    Next r
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim running As Double
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    ws.Cells(1, lastCol + 1).Value = "累計"
    ws.Cells(1, lastCol + 2).Value = "構成比"
    ws.Cells(1, lastCol + 3).Value = "前月比"
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(ws)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsNumeric(ws.Cells(r, valueCol).Value) Then
            ' [ブロック解説: 変数・定数・設定値の宣言]
            ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
            Dim value As Double: value = CDbl(ws.Cells(r, valueCol).Value)
            running = running + value
            ws.Cells(r, lastCol + 1).Value = running
            ws.Cells(r, lastCol + 2).Value = IIf(total <> 0, value / total, 0)
            ws.Cells(r, lastCol + 2).NumberFormat = "0.0%"
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If r > 2 And IsNumeric(ws.Cells(r - 1, valueCol).Value) And CDbl(ws.Cells(r - 1, valueCol).Value) <> 0 Then
                ws.Cells(r, lastCol + 3).Value = (value / CDbl(ws.Cells(r - 1, valueCol).Value)) - 1
                ws.Cells(r, lastCol + 3).NumberFormat = "0.0%"
            End If
        End If
    Next r
End Sub

