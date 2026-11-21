Attribute VB_Name = "Mod173_条件付き書式"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod173_条件付き書式
' - ファイル名: Mod173_条件付き書式.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 条件付き書式達成期限超過上下位設定() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 条件付き書式達成期限超過上下位設定()
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Selection Is Nothing Then Exit Sub
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With Selection.FormatConditions
        .Delete
        .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .Item(.Count).Interior.Color = RGB(255, 199, 206)
        .Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="100"
        .Item(.Count).Interior.Color = RGB(198, 239, 206)
        .AddTop10
        .Item(.Count).TopBottom.TopBottom = xlTop10Top
        .Item(.Count).Interior.Color = RGB(255, 235, 156)
    End With
    MsgBox "条件付き書式を設定しました。", vbInformation
End Sub

