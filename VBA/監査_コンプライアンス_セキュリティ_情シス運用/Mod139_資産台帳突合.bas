Attribute VB_Name = "Mod139_資産台帳突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: ハードウェアPCモバイル資産台帳所在状態確認を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「ハードウェアPCモバイル資産台帳所在状態確認」を選択して実行する。

' ── ハードウェアPCモバイル資産台帳所在状態確認 ──
Public Sub ハードウェアPCモバイル資産台帳所在状態確認()
    Dim inventory As Worksheet: Set inventory = ActiveSheet
    Dim assetName As String: assetName = InputBox("資産台帳シート名", "資産突合")
    Dim asset As Worksheet
    On Error Resume Next
    Set asset = ThisWorkbook.Worksheets(assetName)
    On Error GoTo 0
    If asset Is Nothing Then Exit Sub
    Dim r As Long
    For r = 2 To inventory.Cells(inventory.Rows.Count, 1).End(xlUp).Row
        Dim found As Range
        Set found = asset.Columns(1).Find(CStr(inventory.Cells(r, 1).Value), , xlValues, xlWhole)
        If found Is Nothing Then inventory.Rows(r).Interior.Color = RGB(255, 199, 206)
    Next r
End Sub




