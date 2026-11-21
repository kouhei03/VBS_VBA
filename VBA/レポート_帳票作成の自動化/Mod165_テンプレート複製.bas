Attribute VB_Name = "Mod165_テンプレート複製"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod165_テンプレート複製
' - ファイル名: Mod165_テンプレート複製.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub テンプレートシート複製対象名リネーム() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub テンプレートシート複製対象名リネーム()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim source As Worksheet: Set source = ActiveSheet
    Dim targetName As String: targetName = InputBox("複製後のシート名", "テンプレート複製")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Trim$(targetName) = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim copy As Worksheet
    source.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set copy = ActiveSheet
    copy.Name = targetName
    copy.Cells.Replace What:="{対象名}", Replacement:=targetName, LookAt:=xlPart
    MsgBox "テンプレートを複製しました。", vbInformation
End Sub

