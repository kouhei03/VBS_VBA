Attribute VB_Name = "Mod172_セル表示"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: セル折り返し縮小表示設定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「セル折り返し縮小表示設定」を選択して実行する。

' ── セル折り返し縮小表示設定 ──
Public Sub セル折り返し縮小表示設定()
    Dim mode As String
    mode = InputBox("設定を入力: 1=折り返し、2=縮小表示", "セル表示")
    If mode = "1" Then
        Selection.WrapText = True
    ElseIf mode = "2" Then
        Selection.WrapText = False
        Selection.ShrinkToFit = True
    End If
End Sub




