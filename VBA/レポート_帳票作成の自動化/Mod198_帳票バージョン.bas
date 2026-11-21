Attribute VB_Name = "Mod198_帳票バージョン"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod198_帳票バージョン
' - ファイル名: Mod198_帳票バージョン.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 帳票バージョン作成日時自動印字() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 帳票バージョン作成日時自動印字()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim version As String: version = InputBox("バージョン（例: v1.0）", "帳票バージョン", "v1.0")
    ' [ブロック解説: 一括設定ブロック]
    ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
    With ws.PageSetup
        .RightHeader = "Ver. " & version
        .RightFooter = "作成日時: " & Format$(Now, "yyyy/mm/dd hh:mm")
    End With
    ws.Cells(1, ws.UsedRange.Columns.Count + 1).Value = "Ver." & version
    ws.Cells(2, ws.UsedRange.Columns.Count).Value = "作成: " & Format$(Now, "yyyy/mm/dd hh:mm")
End Sub

