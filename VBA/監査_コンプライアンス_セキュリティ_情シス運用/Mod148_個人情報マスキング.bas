Attribute VB_Name = "Mod148_個人情報マスキング"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 個人情報マスキング匿名化対象件数チェックを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「個人情報マスキング匿名化対象件数チェック」を選択して実行する。

' ── 個人情報マスキング匿名化対象件数チェック ──
Public Sub 個人情報マスキング匿名化対象件数チェック()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim count As Long
    Dim c As Range
    For Each c In ws.UsedRange
        If VarType(c.Value) = vbString And Len(CStr(c.Value)) > 4 Then
            If InStr(CStr(c.Value), "@") > 0 Or c.Column = 2 Then
                c.Value = Left$(CStr(c.Value), 2) & String$(Len(CStr(c.Value)) - 2, "*")
                count = count + 1
            End If
        End If
    Next c
    MsgBox count & " セルをマスキングしました。", vbInformation
End Sub




