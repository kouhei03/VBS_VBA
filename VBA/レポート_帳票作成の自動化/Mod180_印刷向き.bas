Attribute VB_Name = "Mod180_印刷向き"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 印刷の向き自動設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「印刷の向き自動設定」を選択して実行する。

' ── 印刷の向き自動設定 ──
Public Sub 印刷の向き自動設定()
    Dim mode As String: mode = InputBox("向き: 1=縦、2=横", "印刷設定", "2")
    With ActiveSheet.PageSetup
        If mode = "1" Then .Orientation = xlPortrait Else .Orientation = xlLandscape
    End With
End Sub




