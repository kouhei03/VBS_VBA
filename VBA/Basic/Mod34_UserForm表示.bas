Attribute VB_Name = "Mod34_UserForm表示"
Option Explicit
' ファイル処理概要: エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: UserFormを表示を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「UserFormを表示」を選択して実行する。

' ── UserFormを表示 ──
Public Sub UserFormを表示()
    frmBasicSample.Show
End Sub

' ── UserFormを非表示 ──
Public Sub UserFormを非表示()
    On Error Resume Next
    If frmBasicSample.Visible Then frmBasicSample.Hide
    On Error GoTo 0
End Sub




