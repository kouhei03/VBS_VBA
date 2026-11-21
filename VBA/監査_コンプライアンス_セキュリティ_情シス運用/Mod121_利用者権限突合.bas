Attribute VB_Name = "Mod121_利用者権限突合"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 利用者マスタシステム権限一覧突合を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「利用者マスタシステム権限一覧突合」を選択して実行する。

' ── 利用者マスタシステム権限一覧突合 ──
Public Sub 利用者マスタシステム権限一覧突合()
    Dim users As Worksheet: Set users = ActiveSheet
    Dim rightsName As String: rightsName = InputBox("権限一覧シート名", "権限突合")
    Dim rights As Worksheet
    On Error Resume Next
    Set rights = ThisWorkbook.Worksheets(rightsName)
    On Error GoTo 0
    If rights Is Nothing Then Exit Sub
    Dim granted As Object: Set granted = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To rights.Cells(rights.Rows.Count, 1).End(xlUp).Row
        granted(CStr(rights.Cells(r, 1).Value)) = True
    Next r
    Dim output As Worksheet: Set output = GetOrCreateSheet
    output.Cells.Clear
    output.Range("A1:C1").Value = Array("利用者", "権限状態", "備考")
    Dim outRow As Long: outRow = 2
    For r = 2 To users.Cells(users.Rows.Count, 1).End(xlUp).Row
        Dim key As String: key = CStr(users.Cells(r, 1).Value)
        If key <> "" Then
            output.Cells(outRow, 1).Value = key
            output.Cells(outRow, 2).Value = IIf(granted.Exists(key), "権限あり", "権限なし")
            outRow = outRow + 1
        End If
    Next r
End Sub

' ── GetOrCreateSheet ──
Private Function GetOrCreateSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets("利用者権限突合")
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = "利用者権限突合"
    End If
End Function




