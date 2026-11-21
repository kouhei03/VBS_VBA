Attribute VB_Name = "Mod117_繁忙期シミュレーション"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod117_繁忙期シミュレーション
' - ファイル名: Mod117_繁忙期シミュレーション.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 繁忙期閑散期要員計画シミュレーション補助() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 繁忙期閑散期要員計画シミュレーション補助()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim workload As Variant: workload = Application.InputBox("予測業務量", "要員シミュレーション", 100, , , , , 1)
    Dim productivity As Variant: productivity = Application.InputBox("1人あたり処理能力", "要員シミュレーション", 10, , , , , 1)
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If workload = False Or productivity = False Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim required As Double
    required = CDbl(workload) / CDbl(productivity)
    ActiveSheet.Range("A1").Value = "必要要員数"
    ActiveSheet.Range("B1").Value = Application.WorksheetFunction.RoundUp(required, 0)
End Sub

