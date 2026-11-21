Attribute VB_Name = "Mod01_最終行取得"
Option Explicit

Public Function 最終行取得(ByVal ws As Worksheet, Optional ByVal columnNumber As Long = 0) As Long
    Dim found As Range
    If columnNumber > 0 Then
        Set found = ws.Columns(columnNumber).Find(What:="*", LookIn:=xlFormulas, _
                                                  SearchDirection:=xlPrevious)
        If Not found Is Nothing Then 最終行取得 = found.Row
    Else
        Set found = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If Not found Is Nothing Then 最終行取得 = found.Row
    End If
End Function
