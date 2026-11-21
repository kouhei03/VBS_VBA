Attribute VB_Name = "Mod129_Outlook検索"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod129_Outlook検索
' - ファイル名: Mod129_Outlook検索.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub Outlook受信トレイ送信済みアイテム検索() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub Outlook受信トレイ送信済みアイテム検索()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim namespace As Object: Set namespace = outlook.GetNamespace("MAPI")
    Dim keyword As String: keyword = InputBox("検索キーワード", "Outlook検索")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Trim$(keyword) = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "Outlook検索")
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("フォルダ", "件名", "受信日時")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long: r = 2
    SearchFolder namespace.GetDefaultFolder(6), keyword, ws, r
    SearchFolder namespace.GetDefaultFolder(5), keyword, ws, r
    ws.Columns("A:C").AutoFit
End Sub

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Private Sub SearchFolder(ByVal folder As Object, ByVal keyword As String, ByVal ws As Worksheet, ByRef r As Long) の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Private Sub SearchFolder(ByVal folder As Object, ByVal keyword As String, ByVal ws As Worksheet, ByRef r As Long)
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim item As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each item In folder.Items
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If InStr(1, CStr(item.Subject), keyword, vbTextCompare) > 0 Then
            ws.Cells(r, 1).Value = folder.Name
            ws.Cells(r, 2).Value = item.Subject
            ws.Cells(r, 3).Value = item.ReceivedTime
            r = r + 1
        End If
    Next item
End Sub

