Attribute VB_Name = "Mod10_確認ダイアログ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: ユーザー確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ユーザー確認」を選択して実行する。

' ── ユーザー確認 ──
Public Function ユーザー確認(ByVal message As String) As Boolean
    ユーザー確認 = (MsgBox(message, vbQuestion + vbYesNo + vbDefaultButton2, _
                           "確認") = vbYes)
End Function




