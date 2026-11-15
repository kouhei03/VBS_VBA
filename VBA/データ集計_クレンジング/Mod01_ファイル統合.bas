Attribute VB_Name = "Mod01_ファイル統合"
Option Explicit

'==================================================
' Mod01_ファイル統合
' フォルダ内の複数Excelファイル / 複数シートを
' 1つのマスターシートに自動集約する
'
' 前提:
'   - 各ファイル・各シートは1行目がヘッダーの表形式
'   - ヘッダー構成は全ファイルで同一
'==================================================

Private Const MASTER_SHEET As String = "マスター"

'--------------------------------------------------
' メイン: フォルダ内の全Excelファイルを統合
'--------------------------------------------------
Public Sub 複数ファイル統合()
    Dim folderPath As String
    folderPath = SelectFolder("統合するExcelファイルが入ったフォルダを選択")
    If folderPath = "" Then Exit Sub

    Dim wsMaster As Worksheet
    Set wsMaster = GetOrCreateSheet(ThisWorkbook, MASTER_SHEET)
    wsMaster.Cells.Clear

    SpeedUp True
    On Error GoTo Finally

    Dim fileName As String
    Dim wb As Workbook
    Dim fileCount As Long
    Dim rowCount As Long

    fileName = Dir(folderPath & Application.PathSeparator & "*.xls*")
    Do While fileName <> ""
        ' 自分自身・一時ファイルはスキップ
        If fileName <> ThisWorkbook.Name And Left$(fileName, 2) <> "~$" Then
            Set wb = Workbooks.Open( _
                fileName:=folderPath & Application.PathSeparator & fileName, _
                ReadOnly:=True, UpdateLinks:=0)

            Dim ws As Worksheet
            For Each ws In wb.Worksheets
                rowCount = rowCount + CopyToMaster(ws, wsMaster, fileName)
            Next ws

            wb.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If
        fileName = Dir()
    Loop

    wsMaster.Columns.AutoFit

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox fileCount & " 個のファイルから " & rowCount & " 行を統合しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' メイン: このブック内の全シートを統合
'--------------------------------------------------
Public Sub 複数シート統合()
    Dim wsMaster As Worksheet
    Set wsMaster = GetOrCreateSheet(ThisWorkbook, MASTER_SHEET)
    wsMaster.Cells.Clear

    SpeedUp True
    On Error GoTo Finally

    Dim ws As Worksheet
    Dim rowCount As Long
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> MASTER_SHEET Then
            rowCount = rowCount + CopyToMaster(ws, wsMaster, ws.Name)
        End If
    Next ws

    wsMaster.Columns.AutoFit

Finally:
    SpeedUp False
    If Err.Number <> 0 Then
        MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox rowCount & " 行をマスターシートに統合しました。", vbInformation
    End If
End Sub

'--------------------------------------------------
' 1シート分のデータをマスターへコピー
' 戻り値: コピーした行数（ヘッダー除く）
'--------------------------------------------------
Private Function CopyToMaster(ByVal wsSrc As Worksheet, _
                              ByVal wsMaster As Worksheet, _
                              ByVal sourceName As String) As Long
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc)
    lastCol = GetLastCol(wsSrc)
    If lastRow < 2 Then Exit Function   ' ヘッダーのみ / 空シート

    Dim destRow As Long
    destRow = GetLastRow(wsMaster)

    ' マスターが空ならヘッダーをコピーし「取込元」列を追加
    If destRow = 0 Then
        wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Copy _
            wsMaster.Cells(1, 1)
        wsMaster.Cells(1, lastCol + 1).Value = "取込元"
        destRow = 1
    End If

    ' データ部をコピー
    wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Copy
    wsMaster.Cells(destRow + 1, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    ' 取込元を記録
    wsMaster.Range(wsMaster.Cells(destRow + 1, lastCol + 1), _
                   wsMaster.Cells(destRow + lastRow - 1, lastCol + 1)).Value = sourceName

    CopyToMaster = lastRow - 1
End Function
