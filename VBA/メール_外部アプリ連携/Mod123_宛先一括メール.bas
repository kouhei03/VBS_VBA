Attribute VB_Name = "Mod123_宛先一括メール"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: セル指定宛先一括メール作成を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「セル指定宛先一括メール作成」を選択して実行する。

' ── セル指定宛先一括メール作成 ──
Public Sub セル指定宛先一括メール作成()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim mail As Object: Set mail = outlook.CreateItem(0)
        mail.To = CStr(ws.Cells(r, 1).Value)
        mail.CC = CStr(ws.Cells(r, 2).Value)
        mail.BCC = CStr(ws.Cells(r, 3).Value)
        mail.Subject = CStr(ws.Cells(r, 4).Value)
        mail.Body = CStr(ws.Cells(r, 5).Value)
        mail.Save
    Next r
    MsgBox lastRow - 1 & " 件の下書きを作成しました。", vbInformation
End Sub




