Attribute VB_Name = "Mod148_Webページ取得"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod148_Webページ取得
' - ファイル名: Mod148_Webページ取得.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub Webページデータ取得シート転記() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub Webページデータ取得シート転記()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim url As String: url = InputBox("取得するURL", "Webページ取得")
    ' [ブロック解説: 条件分岐]
    ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
    If Trim$(url) = "" Then Exit Sub
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", url, False
    http.Send
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim html As Object: Set html = CreateObject("HTMLFile")
    html.write http.responseText
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "Web取得")
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "URL"
    ws.Cells(1, 2).Value = url
    ws.Cells(2, 1).Value = "本文"
    ws.Cells(2, 2).Value = html.body.innerText
    ws.Cells(3, 1).Value = "状態"
    ws.Cells(3, 2).Value = http.Status
End Sub

