Attribute VB_Name = "Mod56_シート突合"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod56_シート突合
' - ファイル名: Mod56_シート突合.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub シート間キー項目を突合() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub シート間キー項目を突合()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim leftName As String: leftName = InputBox("左側シート名", "キー突合")
    Dim rightName As String: rightName = InputBox("右側シート名", "キー突合")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If leftName = "" Or rightName = "" Then Exit Sub
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Not SheetExists(ThisWorkbook, leftName) Or Not SheetExists(ThisWorkbook, rightName) Then Exit Sub

    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim leftWs As Worksheet: Set leftWs = ThisWorkbook.Worksheets(leftName)
    Dim rightWs As Worksheet: Set rightWs = ThisWorkbook.Worksheets(rightName)
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(rightWs)
        dic(CStr(rightWs.Cells(r, 1).Value)) = True
    Next r
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim out As Worksheet: Set out = GetOrCreateSheet(ThisWorkbook, "突合結果")
    out.Cells.Clear
    out.Cells(1, 1).Value = "キー": out.Cells(1, 2).Value = "結果"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outputRow As Long: outputRow = 2
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 2 To GetLastRow(leftWs)
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim key As String: key = CStr(leftWs.Cells(r, 1).Value)
        out.Cells(outputRow, 1).Value = key
        out.Cells(outputRow, 2).Value = IIf(dic.Exists(key), "一致", "不一致")
        outputRow = outputRow + 1
    Next r
End Sub

