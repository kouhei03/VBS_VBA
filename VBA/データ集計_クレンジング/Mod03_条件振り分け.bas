Attribute VB_Name = "Mod03_条件振り分け"
Option Explicit
' ファイル処理概要: ブック・シートを処理、辞書でキー照合、エラー処理分岐を実行する。
' ファイルないコード処理詳細解説: 条件別シート振り分けを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、マクロ有効ブック、対象シートまたは選択範囲。
' ファイル実行例: ExcelのAlt+F8から「条件別シート振り分け」を選択して実行する。

'==================================================
' Mod03_条件振り分け
' キー列（支店名・担当者名など）の値ごとにデータを
' 抽出し、別シート / 別ブックに自動分割する
'
' 前提: アクティブシートは1行目がヘッダーの表形式
'==================================================

'--------------------------------------------------
' メイン: キー列の値ごとに「別シート」へ振り分け
'--------------------------------------------------
' ── 条件別シート振り分け ──
Public Sub 条件別シート振り分け()
    Dim keyCol As Long
    keyCol = AskKeyColumn()
    If keyCol = 0 Then Exit Sub

    SpeedUp True
    On Error GoTo Finally

    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    Dim keys As Object
    Set keys = CollectKeys(wsSrc, keyCol)

    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc)
    lastCol = GetLastCol(wsSrc)

    Dim key As Variant
    Dim wsDest As Worksheet
    For Each key In keys.keys
        Set wsDest = GetOrCreateSheet(ThisWorkbook, SafeSheetName(CStr(key)))
        wsDest.Cells.Clear
        CopyFiltered wsSrc, keyCol, CStr(key), lastRow, lastCol, wsDest.Range("A1")
        wsDest.Columns.AutoFit
    Next key

    wsSrc.AutoFilterMode = False
    wsSrc.Activate

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox keys.Count & " 件のシートに振り分けました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' メイン: キー列の値ごとに「別ブック」へ振り分け保存
'--------------------------------------------------
' ── 条件別ブック振り分け ──
Public Sub 条件別ブック振り分け()
    Dim keyCol As Long
    keyCol = AskKeyColumn()
    If keyCol = 0 Then Exit Sub

    Dim saveFolder As String
    saveFolder = SelectFolder("分割ブックの保存先フォルダを選択")
    If saveFolder = "" Then Exit Sub

    SpeedUp True
    On Error GoTo Finally

    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    Dim keys As Object
    Set keys = CollectKeys(wsSrc, keyCol)

    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc)
    lastCol = GetLastCol(wsSrc)

    Dim key As Variant
    Dim wbNew As Workbook
    For Each key In keys.keys
        Set wbNew = Workbooks.Add(xlWBATWorksheet)   ' 1シートのみの新規ブック
        CopyFiltered wsSrc, keyCol, CStr(key), lastRow, lastCol, _
                     wbNew.Worksheets(1).Range("A1")
        wbNew.Worksheets(1).Name = SafeSheetName(CStr(key))
        wbNew.Worksheets(1).Columns.AutoFit
        wbNew.SaveAs _
            fileName:=saveFolder & Application.PathSeparator & SafeSheetName(CStr(key)) & ".xlsx", _
            FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next key

    wsSrc.AutoFilterMode = False

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox keys.Count & " 件のブックを保存しました。" & vbCrLf & saveFolder, vbInformation
    End If
End Sub

'--------------------------------------------------
' キー列番号をユーザーに確認
'--------------------------------------------------
' ── AskKeyColumn ──
Private Function AskKeyColumn() As Long
    Dim ans As String
    ans = InputBox("振り分けキーの列番号を入力してください" & vbCrLf & _
                   "（例: A列=1, B列=2 ...）", "キー列の指定", "1")
    If ans = "" Or Not IsNumeric(ans) Then
        AskKeyColumn = 0
    Else
        AskKeyColumn = CLng(ans)
    End If
End Function

'--------------------------------------------------
' キー列のユニーク値を収集（Dictionary）
'--------------------------------------------------
' ── CollectKeys ──
Private Function CollectKeys(ByVal ws As Worksheet, ByVal keyCol As Long) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, r As Long
    Dim v As String
    lastRow = GetLastRow(ws)
    For r = 2 To lastRow
        v = Trim$(CStr(ws.Cells(r, keyCol).Value))
        If v <> "" Then dic(v) = True
    Next r
    Set CollectKeys = dic
End Function

'--------------------------------------------------
' AutoFilterで抽出して貼り付け
'--------------------------------------------------
' ── CopyFiltered ──
Private Sub CopyFiltered(ByVal wsSrc As Worksheet, ByVal keyCol As Long, _
                         ByVal key As String, ByVal lastRow As Long, _
                         ByVal lastCol As Long, ByVal destCell As Range)
    With wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))
        .AutoFilter Field:=keyCol, Criteria1:=key
        .SpecialCells(xlCellTypeVisible).Copy destCell
    End With
    Application.CutCopyMode = False
    wsSrc.AutoFilterMode = False
End Sub




