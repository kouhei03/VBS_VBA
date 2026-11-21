Attribute VB_Name = "Mod25_曜日営業日"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod25_曜日営業日
' - ファイル名: Mod25_曜日営業日.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 曜日判定(ByVal targetDate As Date) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 曜日判定(ByVal targetDate As Date) As String
    曜日判定 = Format$(targetDate, "aaaa")
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 営業日判定(ByVal targetDate As Date, Optional ByVal holidays As Variant) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 営業日判定(ByVal targetDate As Date, Optional ByVal holidays As Variant) As Boolean
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim result As Boolean
    result = (Weekday(targetDate, vbMonday) <= 5)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If result And Not IsMissing(holidays) Then
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If TypeName(holidays) = "Range" Then
            result = (Application.CountIf(holidays, targetDate) = 0)
        End If
    End If
    営業日判定 = result
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 営業日加算(ByVal baseDate As Date, ByVal days As Long) As Date の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 営業日加算(ByVal baseDate As Date, ByVal days As Long) As Date
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim count As Long
    Dim result As Date
    result = baseDate
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While count < days
        result = DateAdd("d", 1, result)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If 営業日判定(result) Then count = count + 1
    Loop
    営業日加算 = result
End Function

