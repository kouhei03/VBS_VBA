Attribute VB_Name = "Mod173_プログレスバー"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: プログレスバー長時間処理可視化を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「プログレスバー長時間処理可視化」を選択して実行する。

' ── プログレスバー長時間処理可視化 ──
Public Sub プログレスバー長時間処理可視化()
    Dim total As Long: total = 100
    Dim i As Long
    Application.StatusBar = "処理を開始します..."
    For i = 1 To total
        Application.StatusBar = "処理中 " & i & "% (" & i & "/" & total & ")"
        DoEvents
    Next i
    Application.StatusBar = False
    MsgBox "処理が完了しました。", vbInformation
End Sub




