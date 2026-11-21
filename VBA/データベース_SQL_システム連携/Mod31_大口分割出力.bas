Attribute VB_Name = "Mod31_大口分割出力"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod31_大口分割出力
' - ファイル名: Mod31_大口分割出力.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub DB抽出大口データシートブック分割() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub DB抽出大口データシートブック分割()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim pageSize As Variant
    pageSize = Application.InputBox("1シートあたりの行数", "分割出力", 50000, , , , , 1)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If pageSize = False Or Not IsNumeric(pageSize) Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim startRow As Long: startRow = 2
    Dim part As Long: part = 1
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While startRow <= lastRow
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim output As Worksheet: Set output = ThisWorkbook.Worksheets.Add
        output.Name = "分割_" & Format$(part, "000")
        ws.Rows(1).Copy output.Rows(1)
        ws.Range(startRow & ":" & WorksheetFunction.Min(startRow + CLng(pageSize) - 1, lastRow)).Copy output.Rows(2)
        startRow = startRow + CLng(pageSize)
        part = part + 1
    Loop
End Sub

