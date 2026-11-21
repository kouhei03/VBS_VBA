Attribute VB_Name = "Mod20_禁止文字"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod20_禁止文字
' - ファイル名: Mod20_禁止文字.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 禁止文字制御文字除去エスケープ() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 禁止文字制御文字除去エスケープ()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim c As Range
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each c In ws.UsedRange
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If VarType(c.Value) = vbString Then
            c.Value = EscapeText(CStr(c.Value))
        End If
    Next c
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function EscapeText(ByVal text As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function EscapeText(ByVal text As String) As String
    text = Replace(text, Chr$(0), "")
    text = Replace(text, Chr$(1), "")
    text = Replace(text, Chr$(13), "")
    text = Replace(text, Chr$(10), "")
    text = Replace(text, Chr$(9), "")
    text = Replace(text, "'", "''")
    text = Replace(text, Chr$(34), """""")
    EscapeText = text
End Function

