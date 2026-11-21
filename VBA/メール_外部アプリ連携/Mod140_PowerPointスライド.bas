Attribute VB_Name = "Mod140_PowerPointスライド"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod140_PowerPointスライド
' - ファイル名: Mod140_PowerPointスライド.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub ExcelデータからPowerPointスライド自動生成() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub ExcelデータからPowerPointスライド自動生成()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    ppt.Visible = True
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim presentation As Object: Set presentation = ppt.Presentations.Add
    Dim r As Long
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For r = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' [ブロック解説: 変数・定数・設定値の宣言]
        ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
        Dim slide As Object: Set slide = presentation.Slides.Add(presentation.Slides.Count + 1, 1)
        slide.Shapes(1).TextFrame.TextRange.Text = CStr(ws.Cells(r, 1).Value)
        ' [ブロック解説: 条件分岐]
        ' 指定した条件が真のときだけ処理を実行します。比較対象、空欄判定、型変換、否定条件を確認してください。
        If ws.Cells(r, 2).Value <> "" Then
            slide.Shapes(2).TextFrame.TextRange.Text = CStr(ws.Cells(r, 2).Value)
        End If
    Next r
    MsgBox "スライドを生成しました。", vbInformation
End Sub

