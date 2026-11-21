Attribute VB_Name = "Mod183_共通ユーティリティ"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 次営業日を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「次営業日」を選択して実行する。

' ── 次営業日 ──
Public Function 次営業日(ByVal baseDate As Variant) As Date
    Dim d As Date: d = CDate(baseDate) + 1
    Do While Weekday(d, vbMonday) > 5
        d = d + 1
    Loop
    次営業日 = d
End Function

' ── 共通処理ログ出力 ──
Public Sub 共通処理ログ出力(ByVal message As String)
    Dim filePath As String: filePath = ThisWorkbook.Path & "\common-tool.log"
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open filePath For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy/mm/dd hh:mm:ss") & vbTab & message
    Close #fileNumber
End Sub




