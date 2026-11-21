Attribute VB_Name = "Mod136_有効期限一覧"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 証明書ドメインライセンス契約期限一覧アラートを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「証明書ドメインライセンス契約期限一覧アラート」を選択して実行する。

' ── 証明書ドメインライセンス契約期限一覧アラート ──
Public Sub 証明書ドメインライセンス契約期限一覧アラート()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim threshold As Variant: threshold = Application.InputBox("期限何日前から警告", "有効期限", 30, , , , , 1)
    If threshold = False Or Not IsNumeric(threshold) Then Exit Sub
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(r, 2).Value) Then
            Dim remaining As Long: remaining = DateDiff("d", Date, CDate(ws.Cells(r, 2).Value))
            ws.Cells(r, 3).Value = remaining
            If remaining < 0 Then
                ws.Rows(r).Interior.Color = RGB(255, 199, 206)
            ElseIf remaining <= CLng(threshold) Then
                ws.Rows(r).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    Next r
End Sub




