Attribute VB_Name = "Mod198_権限表示制御"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 権限ロール画面表示実行ボタン制御を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「権限ロール画面表示実行ボタン制御」を選択して実行する。

' ── 権限ロール画面表示実行ボタン制御 ──
Public Sub 権限ロール画面表示実行ボタン制御()
    Dim role As String: role = InputBox("ロールを入力（管理者 / 一般）", "権限制御")
    Dim c As Shape
    For Each c In ActiveSheet.Shapes
        c.Visible = IIf(role = "管理者", msoTrue, msoFalse)
    Next c
    If role = "一般" Then MsgBox "一般ユーザー向け表示に切り替えました。", vbInformation
End Sub




