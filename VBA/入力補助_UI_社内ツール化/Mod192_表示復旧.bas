Attribute VB_Name = "Mod192_表示復旧"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 計算モード画面表示復旧ボタンを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「計算モード画面表示復旧ボタン」を選択して実行する。

' ── 計算モード画面表示復旧ボタン ──
Public Sub 計算モード画面表示復旧ボタン()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    ActiveWindow.Zoom = 100
    MsgBox "画面表示と計算モードを復旧しました。", vbInformation
End Sub




