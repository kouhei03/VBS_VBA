Attribute VB_Name = "Mod103_営業日計算"
Option Explicit

Public Sub 営業日カレンダー期日納期締切計算()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim baseDate As Variant: baseDate = Application.InputBox("基準日", "営業日計算", Date, , , , , 1)
    Dim addDays As Variant: addDays = Application.InputBox("加算営業日数", "営業日計算", 5, , , , , 1)
    If baseDate = False Or addDays = False Then Exit Sub
    Dim dateValue As Date: dateValue = CDate(baseDate)
    Dim count As Long
    Do While count < CLng(addDays)
        dateValue = dateValue + 1
        If Weekday(dateValue, vbMonday) <= 5 Then count = count + 1
    Loop
    ws.Range("A1").Value = "計算期日"
    ws.Range("B1").Value = dateValue
    ws.Range("B1").NumberFormat = "yyyy/mm/dd"
End Sub
