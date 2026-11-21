Attribute VB_Name = "Mod181_印刷縮尺"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 印刷拡大縮小横1縦自動設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「印刷拡大縮小横1縦自動設定」を選択して実行する。

' ── 印刷拡大縮小横1縦自動設定 ──
Public Sub 印刷拡大縮小横1縦自動設定()
    Dim mode As String: mode = InputBox("設定: 1=横1、2=縦自動、3=縮尺指定", "印刷縮尺")
    With ActiveSheet.PageSetup
        If mode = "1" Then
            .FitToPagesWide = 1
            .FitToPagesTall = 0
        ElseIf mode = "2" Then
            .FitToPagesWide = 0
            .FitToPagesTall = 1
        ElseIf mode = "3" Then
            .Zoom = Application.InputBox("縮尺率（%）", "印刷縮尺", 100, , , , , 1)
        End If
    End With
End Sub




