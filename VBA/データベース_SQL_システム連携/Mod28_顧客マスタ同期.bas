Attribute VB_Name = "Mod28_顧客マスタ同期"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 顧客マスタ新規変更廃止同期リストを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「顧客マスタ新規変更廃止同期リスト」を選択して実行する。

' ── 顧客マスタ新規変更廃止同期リスト ──
Public Sub 顧客マスタ新規変更廃止同期リスト()
    Dim masterWs As Worksheet: Set masterWs = ActiveSheet
    Dim compareName As String: compareName = InputBox("比較先シート名", "顧客同期")
    Dim compareWs As Worksheet
    On Error Resume Next
    Set compareWs = ThisWorkbook.Worksheets(compareName)
    On Error GoTo 0
    If compareWs Is Nothing Then Exit Sub
    Dim current As Object: Set current = CreateObject("Scripting.Dictionary")
    Dim previous As Object: Set previous = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To masterWs.Cells(masterWs.Rows.Count, 1).End(xlUp).Row
        current(CStr(masterWs.Cells(r, 1).Value)) = CStr(masterWs.Cells(r, 2).Value)
    Next r
    For r = 2 To compareWs.Cells(compareWs.Rows.Count, 1).End(xlUp).Row
        previous(CStr(compareWs.Cells(r, 1).Value)) = CStr(compareWs.Cells(r, 2).Value)
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("顧客コード", "区分", "値")
    Dim outRow As Long: outRow = 2
    Dim key As Variant
    For Each key In current.Keys
        If Not previous.Exists(CStr(key)) Then
            output.Cells(outRow, 1).Value = key
            output.Cells(outRow, 2).Value = "新規"
            outRow = outRow + 1
        ElseIf CStr(current(key)) <> CStr(previous(key)) Then
            output.Cells(outRow, 1).Value = key
            output.Cells(outRow, 2).Value = "変更"
            outRow = outRow + 1
        End If
    Next key
    For Each key In previous.Keys
        If Not current.Exists(CStr(key)) Then
            output.Cells(outRow, 1).Value = key
            output.Cells(outRow, 2).Value = "廃止"
            outRow = outRow + 1
        End If
    Next key
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("顧客マスタ同期")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "顧客マスタ同期"
    End If
End Function




