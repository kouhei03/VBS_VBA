Attribute VB_Name = "Mod125_セル差し込み"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod125_セル差し込み
' - ファイル名: Mod125_セル差し込み.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 件名本文セル値差し込み() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 件名本文セル値差し込み()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim subjectTemplate As String: subjectTemplate = InputBox("件名テンプレート（{A2}形式）", "差し込み")
    Dim bodyTemplate As String: bodyTemplate = InputBox("本文テンプレート（{A2}形式）", "差し込み")
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.Subject = ReplaceCellTokens(subjectTemplate)
    mail.Body = ReplaceCellTokens(bodyTemplate)
    mail.Save
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Function ReplaceCellTokens(ByVal text As String) As String の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Function ReplaceCellTokens(ByVal text As String) As String
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{([A-Z]+\d+)\}"
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim matches As Object: Set matches = re.Execute(text)
    Dim match As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each match In matches
        text = Replace(text, match.Value, CStr(ActiveSheet.Range(match.SubMatches(0)).Value))
    Next match
    ReplaceCellTokens = text
End Function

