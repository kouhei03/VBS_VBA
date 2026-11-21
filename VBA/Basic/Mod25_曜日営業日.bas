Attribute VB_Name = "Mod25_曜日営業日"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 曜日判定を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「曜日判定」を選択して実行する。

' ── 曜日判定 ──
Public Function 曜日判定(ByVal targetDate As Date) As String
    曜日判定 = Format$(targetDate, "aaaa")
End Function

' ── 営業日判定 ──
Public Function 営業日判定(ByVal targetDate As Date, Optional ByVal holidays As Variant) As Boolean
    Dim result As Boolean
    result = (Weekday(targetDate, vbMonday) <= 5)
    If result And Not IsMissing(holidays) Then
        If TypeName(holidays) = "Range" Then
            result = (Application.CountIf(holidays, targetDate) = 0)
        End If
    End If
    営業日判定 = result
End Function

' ── 営業日加算 ──
Public Function 営業日加算(ByVal baseDate As Date, ByVal days As Long) As Date
    Dim count As Long
    Dim result As Date
    result = baseDate
    Do While count < days
        result = DateAdd("d", 1, result)
        If 営業日判定(result) Then count = count + 1
    Loop
    営業日加算 = result
End Function




