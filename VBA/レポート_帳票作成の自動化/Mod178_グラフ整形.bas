Attribute VB_Name = "Mod178_グラフ整形"
Option Explicit
' [設定解説追加済み]
' [ブロック解説: モジュール設定]
' - モジュール名: Mod178_グラフ整形
' - ファイル名: Mod178_グラフ整形.bas
' - 目的: このモジュールに定義されたマクロまたは関数を、Excelの標準オブジェクトを使って実行します。
' - 実行前設定: 対象ブック、対象シート、選択範囲、入力ダイアログの値が正しいことを確認してください。
' - 出力設定: シート・ブック・ファイルへの出力先を、既存データを上書きしない運用か確認してください。
' - セキュリティ設定: 外部アプリ、ネットワーク、DB、レジストリ、メール送信へ影響する場合はテスト環境で先に確認します。
' - 注意: 処理前にバックアップを作成し、確認ダイアログ、権限、保存先、エラー処理の設定を確認してください。

' [ブロック解説: プロシージャ]
' 宣言ブロックです。Public Sub グラフ位置サイズタイトル凡例自動整形() の処理範囲、引数、対象ブック・シート、利用する設定値を確認してから実行します。
Public Sub グラフ位置サイズタイトル凡例自動整形()
    ' [ブロック解説: 変数・定数・設定値の宣言]
    ' このブロックで型、初期値、対象範囲、オブジェクト参照などを定義します。実行前に型と初期値を確認してください。
    Dim titleText As String: titleText = InputBox("グラフタイトル", "グラフ整形")
    Dim chartObject As ChartObject
    ' [ブロック解説: 繰り返し処理]
    ' 対象コレクションまたは範囲を順番に処理します。開始値、終了値、ループ変数、削除対象・書き込み対象を確認してください。
    For Each chartObject In ActiveSheet.ChartObjects
        ' [ブロック解説: 一括設定ブロック]
        ' Withブロック内のオブジェクトへ、セル形式、印刷設定、オプションなどをまとめて設定します。設定対象と上書き範囲に注意します。
        With chartObject
            .Left = 50
            .Top = 120
            .Width = 600
            .Height = 360
            .Chart.HasTitle = True
            .Chart.ChartTitle.Text = titleText
            .Chart.HasLegend = True
            .Chart.Legend.Position = xlLegendPositionBottom
        End With
    Next chartObject
End Sub

