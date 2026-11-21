Attribute VB_Name = "Mod65_異常値抽出"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod65_異常値抽出
' - ファイル名: Mod65_異常値抽出.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 異常値を抽出してフラグ付与() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 異常値を抽出してフラグ付与()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim col As Long: col = CLng(InputBox("対象列番号", "異常値検出", "3"))
    Dim threshold As Double: threshold = CDbl(InputBox("閾値", "異常値検出", "100"))
    Dim lastCol As Long: lastCol = Application.WorksheetFunction.Max(GetLastCol(ws), col + 1)
    ws.Cells(1, lastCol).Value = "異常値フラグ"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(ws)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsNumeric(ws.Cells(r, col).Value) Then
            ws.Cells(r, lastCol).Value = IIf(CDbl(ws.Cells(r, col).Value) > threshold, "異常", "")
            ' [ブロック解説: 条件分岐]
            ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
            If CDbl(ws.Cells(r, col).Value) > threshold Then
                ws.Cells(r, col).Interior.Color = RGB(255, 199, 206)
            End If
        End If
    Next r
End Sub

