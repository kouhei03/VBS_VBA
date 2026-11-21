Attribute VB_Name = "Mod125_セル差し込み"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: 件名本文セル値差し込みを中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「件名本文セル値差し込み」を選択して実行する。

' ── 件名本文セル値差し込み ──
Public Sub 件名本文セル値差し込み()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim subjectTemplate As String: subjectTemplate = InputBox("件名テンプレート（{A2}形式）", "差し込み")
    Dim bodyTemplate As String: bodyTemplate = InputBox("本文テンプレート（{A2}形式）", "差し込み")
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.Subject = ReplaceCellTokens(subjectTemplate)
    mail.Body = ReplaceCellTokens(bodyTemplate)
    mail.Save
End Sub

' ── ReplaceCellTokens ──
Private Function ReplaceCellTokens(ByVal text As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{([A-Z]+\d+)\}"
    Dim matches As Object: Set matches = re.Execute(text)
    Dim match As Object
    For Each match In matches
        text = Replace(text, match.Value, CStr(ActiveSheet.Range(match.SubMatches(0)).Value))
    Next match
    ReplaceCellTokens = text
End Function




