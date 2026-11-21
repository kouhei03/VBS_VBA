Attribute VB_Name = "Mod132_本文定型抽出"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod132_本文定型抽出
' - ファイル名: Mod132_本文定型抽出.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub 受信本文定型データ抽出シート転記() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub 受信本文定型データ抽出シート転記()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim inbox As Object: Set inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(ThisWorkbook, "受信本文抽出")
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("件名", "受信日時", "送信元", "本文")
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim r As Long: r = 2
    Dim item As Object
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each item In inbox.Items
        ws.Cells(r, 1).Value = item.Subject
        ws.Cells(r, 2).Value = item.ReceivedTime
        ws.Cells(r, 3).Value = item.SenderName
        ws.Cells(r, 4).Value = item.Body
        r = r + 1
    Next item
    ws.Columns("A:D").AutoFit
End Sub

