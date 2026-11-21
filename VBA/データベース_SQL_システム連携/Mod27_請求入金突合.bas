Attribute VB_Name = "Mod27_請求入金突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 請求入金消込突合差異一覧を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「請求入金消込突合差異一覧」を選択して実行する。

' ── 請求入金消込突合差異一覧 ──
Public Sub 請求入金消込突合差異一覧()
    Dim billWs As Worksheet: Set billWs = ActiveSheet
    Dim paymentName As String: paymentName = InputBox("入金シート名", "請求入金突合")
    Dim paymentWs As Worksheet
    On Error Resume Next
    Set paymentWs = ThisWorkbook.Worksheets(paymentName)
    On Error GoTo 0
    If paymentWs Is Nothing Then Exit Sub
    Dim payments As Object: Set payments = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To paymentWs.Cells(paymentWs.Rows.Count, 1).End(xlUp).Row
        payments(CStr(paymentWs.Cells(r, 1).Value)) = CDbl(paymentWs.Cells(r, 2).Value)
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:D1").Value = Array("請求No", "請求額", "入金額", "差額")
    Dim outRow As Long: outRow = 2
    For r = 2 To billWs.Cells(billWs.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(billWs.Cells(r, 1).Value)
        Dim billed As Double: billed = CDbl(billWs.Cells(r, 2).Value)
        Dim paid As Double: paid = 0
        If payments.Exists(key) Then paid = CDbl(payments(key))
        If billed <> paid Then
            output.Cells(outRow, 1).Value = key
            output.Cells(outRow, 2).Value = billed
            output.Cells(outRow, 3).Value = paid
            output.Cells(outRow, 4).Value = billed - paid
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("請求入金差異")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "請求入金差異"
    End If
End Function




