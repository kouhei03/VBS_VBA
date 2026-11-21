Attribute VB_Name = "Mod26_受注出荷照合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 受注出荷実績照合未出荷リストを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「受注出荷実績照合未出荷リスト」を選択して実行する。

' ── 受注出荷実績照合未出荷リスト ──
Public Sub 受注出荷実績照合未出荷リスト()
    Dim orderWs As Worksheet: Set orderWs = ActiveSheet
    Dim shipmentSheetName As String: shipmentSheetName = InputBox("出荷シート名", "出荷照合")
    Dim shipmentWs As Worksheet
    On Error Resume Next
    Set shipmentWs = ThisWorkbook.Worksheets(shipmentSheetName)
    On Error GoTo 0
    If shipmentWs Is Nothing Then Exit Sub
    Dim shipped As Object: Set shipped = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To shipmentWs.Cells(shipmentWs.Rows.Count, 1).End(xlUp).Row
        shipped(CStr(shipmentWs.Cells(r, 1).Value)) = True
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    Dim outRow As Long: outRow = 1
    For r = 2 To orderWs.Cells(orderWs.Rows.Count, 1).End(xlUp).Row
        If Not shipped.Exists(CStr(orderWs.Cells(r, 1).Value)) Then
            orderWs.Rows(r).Copy output.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("未出荷リスト")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "未出荷リスト"
    End If
End Function




