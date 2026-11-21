Attribute VB_Name = "Mod183_共通ユーティリティ"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod183_共通ユーティリティ
' - ファイル名: Mod183_共通ユーティリティ.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Function 次営業日(ByVal baseDate As Variant) As Date の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Function 次営業日(ByVal baseDate As Variant) As Date
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim d As Date: d = CDate(baseDate) + 1
    ' [ブロック解説: 条件付き繰り返し]
    ' 終了条件が満たされるまで処理を繰り返します。無限ループ防止のためカウンターと終了条件を確認してください。
    Do While Weekday(d, vbMonday) > 5
        d = d + 1
    Loop
    次営業日 = d
End Function

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 共通処理ログ出力(ByVal message As String) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 共通処理ログ出力(ByVal message As String)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim filePath As String: filePath = ThisWorkbook.Path & "\common-tool.log"
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open filePath For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy/mm/dd hh:mm:ss") & vbTab & message
    Close #fileNumber
End Sub

