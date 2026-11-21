Attribute VB_Name = "Mod190_配布ブック"
Option Explicit

Public Sub 配布用ブック数式値化不要シート削除入力欄ロック()
    If MsgBox("配布用ブックを作成します。元ブックは変更しませんか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Dim copyWb As Workbook: Set copyWb = Workbooks.Add(1)
    ThisWorkbook.ActiveSheet.Copy Before:=copyWb.Worksheets(1)
    copyWb.Worksheets(1).UsedRange.Value = copyWb.Worksheets(1).UsedRange.Value
    copyWb.Worksheets(copyWb.Worksheets.Count).Delete
    copyWb.Worksheets(1).Cells.Locked = False
    copyWb.Worksheets(1).UsedRange.Locked = True
    copyWb.Worksheets(1).Protect Password:="dist"
    Dim path As String
    path = Application.GetSaveAsFilename("配布用ブック.xlsx", "Excelブック (*.xlsx),*.xlsx")
    If path <> "False" Then copyWb.SaveAs path, xlOpenXMLWorkbook
End Sub
