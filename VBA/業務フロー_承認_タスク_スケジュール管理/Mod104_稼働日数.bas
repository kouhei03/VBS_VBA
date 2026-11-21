Attribute VB_Name = "Mod104_稼働日数"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 祝日年末年始GW稼働日数算出を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「祝日年末年始GW稼働日数算出」を選択して実行する。

' ── 祝日年末年始GW稼働日数算出 ──
Public Sub 祝日年末年始GW稼働日数算出()
    Dim startDate As Variant: startDate = Application.InputBox("開始日", "稼働日数")
    Dim endDate As Variant: endDate = Application.InputBox("終了日", "稼働日数")
    If startDate = False Or endDate = False Then Exit Sub
    Dim count As Long: count = 0
    Dim d As Date
    For d = CDate(startDate) To CDate(endDate)
        If Weekday(d, vbMonday) <= 5 And Not IsJapaneseHoliday(d) Then count = count + 1
    Next d
    MsgBox count & " 営業日です。", vbInformation
End Sub

' ── IsJapaneseHoliday ──
Private Function IsJapaneseHoliday(ByVal d As Date) As Boolean
    Select Case Format$(d, "mmdd")
        Case "0101", "0102", "0103", "0429", "0503", "0504", "0505", "0811", "1113", "1123"
            IsJapaneseHoliday = True
    End Select
End Function




