Attribute VB_Name = "Mod182_印刷タイトル"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 印刷タイトル行列固定設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「印刷タイトル行列固定設定」を選択して実行する。

' ── 印刷タイトル行列固定設定 ──
Public Sub 印刷タイトル行列固定設定()
    Dim repeatRows As String: repeatRows = InputBox("繰り返し行（例: $1:$2、空欄可）", "印刷タイトル")
    Dim repeatColumns As String: repeatColumns = InputBox("繰り返し列（例: $A:$B、空欄可）", "印刷タイトル")
    With ActiveSheet.PageSetup
        If repeatRows <> "" Then .PrintTitleRows = repeatRows
        If repeatColumns <> "" Then .PrintTitleColumns = repeatColumns
    End With
End Sub




