Attribute VB_Name = "Mod208_クリーンアップ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: テストクリーンアップ初期状態復元を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「テストクリーンアップ初期状態復元」を選択して実行する。

' ── テストクリーンアップ初期状態復元 ──
Public Sub テストクリーンアップ初期状態復元()
    Dim ws As Worksheet: Set ws = ActiveSheet
    If MsgBox("テスト結果と挿入データを初期化しますか？", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    ws.UsedRange.ClearContents
    ws.UsedRange.ClearFormats
    MsgBox "テストシートを初期化しました。", vbInformation
End Sub




