Attribute VB_Name = "Mod40_完了通知"
Option Explicit
' ファイル処理概要: エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 処理完了通知を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「処理完了通知」を選択して実行する。

' ── 処理完了通知 ──
Public Sub 処理完了通知(Optional ByVal message As String = "処理が完了しました。")
    On Error Resume Next
    Beep
    On Error GoTo 0
    MsgBox message, vbInformation, "完了"
End Sub




