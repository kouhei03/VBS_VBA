Attribute VB_Name = "Mod00_共通"
Option Explicit

'==================================================
' Mod00_共通
' 各モジュールから利用する共通ヘルパー関数
'==================================================

'--------------------------------------------------
' 指定シートの最終行を取得（全列を対象）
'--------------------------------------------------
Public Function GetLastRow(ByVal ws As Worksheet) As Long
    Dim rng As Range
    Set rng = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If rng Is Nothing Then
        GetLastRow = 0
    Else
        GetLastRow = rng.Row
    End If
End Function

'--------------------------------------------------
' 指定シートの最終列を取得（全行を対象）
'--------------------------------------------------
Public Function GetLastCol(ByVal ws As Worksheet) As Long
    Dim rng As Range
    Set rng = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If rng Is Nothing Then
        GetLastCol = 0
    Else
        GetLastCol = rng.Column
    End If
End Function

'--------------------------------------------------
' シートの存在確認
'--------------------------------------------------
Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'--------------------------------------------------
' シートを取得（なければ新規作成）
'--------------------------------------------------
Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If SheetExists(wb, sheetName) Then
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

'--------------------------------------------------
' シート名として使えない文字を置換
'--------------------------------------------------
Public Function SafeSheetName(ByVal s As String) As String
    Dim badChars As Variant
    Dim i As Long
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "_")
    Next i
    If Len(s) > 31 Then s = Left$(s, 31)
    If Len(s) = 0 Then s = "_blank_"
    SafeSheetName = s
End Function

'--------------------------------------------------
' フォルダ選択ダイアログ
'--------------------------------------------------
Public Function SelectFolder(Optional ByVal title As String = "フォルダを選択してください") As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""   ' キャンセル時は空文字
        End If
    End With
End Function

'--------------------------------------------------
' 高速化設定の ON/OFF
'--------------------------------------------------
Public Sub SpeedUp(ByVal enable As Boolean)
    With Application
        .ScreenUpdating = Not enable
        .EnableEvents = Not enable
        .Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
        .DisplayAlerts = Not enable
    End With
End Sub
