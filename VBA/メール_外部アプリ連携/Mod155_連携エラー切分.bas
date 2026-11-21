Attribute VB_Name = "Mod155_連携エラー切分"
Option Explicit
' ファイル処理概要: Wordを操作、PowerPointを操作、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 複数アプリ連携エラー切り分けクローズ処理を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「複数アプリ連携エラー切り分けクローズ処理」を選択して実行する。

' ── 複数アプリ連携エラー切り分けクローズ処理 ──
Public Sub 複数アプリ連携エラー切り分けクローズ処理()
    On Error GoTo Failed
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim word As Object: Set word = CreateObject("Word.Application")
    Dim ppt As Object: Set ppt = CreateObject("PowerPoint.Application")
    MsgBox "Outlook、Word、PowerPointへの接続に成功しました。", vbInformation
CleanExit:
    On Error Resume Next
    If Not word Is Nothing Then word.Quit
    If Not ppt Is Nothing Then ppt.Quit
    Exit Sub
Failed:
    MsgBox "連携エラー: " & Err.Description & " (番号 " & Err.Number & ")", vbCritical
    Resume CleanExit
End Sub




