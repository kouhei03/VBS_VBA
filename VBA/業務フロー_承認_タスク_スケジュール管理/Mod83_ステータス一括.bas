Attribute VB_Name = "Mod83_ステータス一括"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ステータス一括更新を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ステータス一括更新」を選択して実行する。

' ── ステータス一括更新 ──
Public Sub ステータス一括更新()
    Dim status As String
    status = InputBox("更新ステータス（未着手 / 進行中 / 完了 / 保留）", "ステータス更新")
    If Selection Is Nothing Then Exit Sub
    Selection.Value = status
End Sub




