Attribute VB_Name = "Mod08_エラー処理"
Option Explicit
' ファイル処理概要: エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 安全な処理例を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「安全な処理例」を選択して実行する。

' ── 安全な処理例 ──
Public Sub 安全な処理例()
    On Error GoTo Failed

    ' ここに処理を記述
    Dim value As Long
    value = CLng(ActiveSheet.Range("A1").Value)

CleanExit:
    Exit Sub
Failed:
    MsgBox "エラー " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub




