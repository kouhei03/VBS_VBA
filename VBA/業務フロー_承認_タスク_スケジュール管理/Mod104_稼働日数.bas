Attribute VB_Name = "Mod104_稼働日数"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod104_稼働日数
' - ファイル名: Mod104_稼働日数.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 祝日年末年始GW稼働日数算出() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 祝日年末年始GW稼働日数算出()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim startDate As Variant: startDate = Application.InputBox("開始日", "稼働日数")
    Dim endDate As Variant: endDate = Application.InputBox("終了日", "稼働日数")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If startDate = False Or endDate = False Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim count As Long: count = 0
    Dim d As Date
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For d = CDate(startDate) To CDate(endDate)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If Weekday(d, vbMonday) <= 5 And Not IsJapaneseHoliday(d) Then count = count + 1
    Next d
    MsgBox count & " 営業日です。", vbInformation
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function IsJapaneseHoliday(ByVal d As Date) As Boolean の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function IsJapaneseHoliday(ByVal d As Date) As Boolean
    ' [ブロック解説: 分岐条件の設定]
    ' 入力値やステータスに応じた分岐を設定します。未定義値の既定動作、すべきでない処理へのフォールバックを確認してください。
    Select Case Format$(d, "mmdd")
        Case "0101", "0102", "0103", "0429", "0503", "0504", "0505", "0811", "1113", "1123"
            IsJapaneseHoliday = True
    End Select
End Function

