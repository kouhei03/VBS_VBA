Attribute VB_Name = "Mod75_ランダム抽出"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod75_ランダム抽出
' - ファイル名: Mod75_ランダム抽出.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ランダムサンプリング抽出() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ランダムサンプリング抽出()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim sampleSize As Long: sampleSize = CLng(InputBox("抽出件数", "ランダム抽出", "10"))
    Dim out As Worksheet: Set out = GetOrCreateSheet(ws.Parent, "ランダムサンプル")
    out.Cells.Clear
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long: lastRow = GetLastRow(ws)
    Dim lastCol As Long: lastCol = GetLastCol(ws)
    out.Range(out.Cells(1, 1), out.Cells(1, lastCol)).Value = _
        ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Value
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Randomize
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outputRow As Long: outputRow = 2
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While outputRow < sampleSize + 2 And dic.Count < lastRow - 1
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim r As Long: r = Int((lastRow - 1) * Rnd) + 2
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Not dic.Exists(r) Then
            dic.Add r, True
            out.Rows(outputRow).Value = ws.Rows(r).Value
            outputRow = outputRow + 1
        End If
    Loop
End Sub

