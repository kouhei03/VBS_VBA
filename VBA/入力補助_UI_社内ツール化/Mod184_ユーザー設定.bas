Attribute VB_Name = "Mod184_ユーザー設定"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod184_ユーザー設定
' - ファイル名: Mod184_ユーザー設定.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ユーザー部署別設定保存復元() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ユーザー部署別設定保存復元()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim key As String: key = CreateObject("WScript.Network").UserName
    Dim setting As String: setting = InputBox("保存または復元する設定値", "ユーザー設定")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If setting = "" Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim restored As String
        ' [ブロック解説: エラー処理設定]
        ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
        On Error Resume Next
        restored = CStr(CreateObject("WScript.Shell").RegRead("HKCU\Software\VBAInternalTools\" & key))
        ' [ブロック解説: エラー処理設定]
        ' エラー発生時の処理分岐を定義します。エラー番号、エラー内容、復旧処理、画面更新の復帰を確認してください。
        On Error GoTo 0
        MsgBox restored, vbInformation
    Else
        CreateObject("WScript.Shell").RegWrite "HKCU\Software\VBAInternalTools\" & key, setting, "REG_SZ"
        MsgBox "設定を保存しました。", vbInformation
    End If
End Sub

