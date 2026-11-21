Attribute VB_Name = "Mod02_最終列取得"
Option Explicit

Public Function 最終列取得(ByVal ws As Worksheet, Optional ByVal rowNumber As Long = 0) As Long
    Dim found As Range
    If rowNumber > 0 Then
        Set found = ws.Rows(rowNumber).Find(What:="*", LookIn:=xlFormulas, _
                                            SearchDirection:=xlPrevious)
        If Not found Is Nothing Then 最終列取得 = found.Column
    Else
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
        If Not found Is Nothing Then 最終列取得 = found.Column
    End If
End Function
