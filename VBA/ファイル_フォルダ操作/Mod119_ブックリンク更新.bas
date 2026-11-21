Attribute VB_Name = "Mod119_ブックリンク更新"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod119_ブックリンク更新
' - ファイル名: Mod119_ブックリンク更新.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 複数ブックリンク更新制御() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 複数ブックリンク更新制御()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim mode As String
    mode = InputBox("リンク更新を制御: 1=更新する、2=更新しない、3=状態確認", "リンク更新", "2")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If mode = "1" Then
        Application.AskToUpdateLinks = True
        ThisWorkbook.UpdateLinks = xlUpdateLinksAlways
    ElseIf mode = "2" Then
        Application.AskToUpdateLinks = False
        ThisWorkbook.UpdateLinks = xlUpdateLinksNever
    ElseIf mode = "3" Then
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim links As Variant
        links = ThisWorkbook.LinkSources(xlExcelLinks)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If IsEmpty(links) Then
            MsgBox "外部リンクはありません。", vbInformation
        Else
            MsgBox "ブックの外部リンク数: " & UBound(links) - LBound(links) + 1, vbInformation
        End If
    End If
End Sub

