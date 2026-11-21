Attribute VB_Name = "Mod103_ブック形式変換"
Option Explicit

Public Sub ブック形式変換保存()
    Dim mode As String
    mode = InputBox("保存形式を入力: 1=xlsx、2=xlsm、3=xls、4=xlsb", "形式変換", "1")
    Dim destination As Long
    Select Case mode
        Case "1": destination = xlOpenXMLWorkbook
        Case "2": destination = xlOpenXMLWorkbookMacroEnabled
        Case "3": destination = xlExcel8
        Case "4": destination = xlExcel12
        Case Else: Exit Sub
    End Select
    Dim baseName As String
    baseName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    ThisWorkbook.SaveAs baseName & "_" & Format$(Now, "yyyymmdd") & GetExtension(mode), destination
End Sub

Private Function GetExtension(ByVal mode As String) As String
    Select Case mode
        Case "1": GetExtension = ".xlsx"
        Case "2": GetExtension = ".xlsm"
        Case "3": GetExtension = ".xls"
        Case "4": GetExtension = ".xlsb"
    End Select
End Function
