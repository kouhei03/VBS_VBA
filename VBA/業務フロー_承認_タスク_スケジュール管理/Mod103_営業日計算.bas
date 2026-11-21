Attribute VB_Name = "Mod103_営業日計算"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod103_営業日計算
' - ファイル名: Mod103_営業日計算.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 営業日カレンダー期日納期締切計算() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 営業日カレンダー期日納期締切計算()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim baseDate As Variant: baseDate = Application.InputBox("基準日", "営業日計算", Date, , , , , 1)
    Dim addDays As Variant: addDays = Application.InputBox("加算営業日数", "営業日計算", 5, , , , , 1)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If baseDate = False Or addDays = False Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim dateValue As Date: dateValue = CDate(baseDate)
    Dim count As Long
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While count < CLng(addDays)
        dateValue = dateValue + 1
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Weekday(dateValue, vbMonday) <= 5 Then count = count + 1
    Loop
    ws.Range("A1").Value = "計算期日"
    ws.Range("B1").Value = dateValue
    ws.Range("B1").NumberFormat = "yyyy/mm/dd"
End Sub

