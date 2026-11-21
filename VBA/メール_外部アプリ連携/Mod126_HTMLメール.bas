Attribute VB_Name = "Mod126_HTMLメール"
Option Explicit
' ファイル処理概要: セル・シート・ブック・ダイアログを使用した業務処理を実行する。
' ファイルないコード処理詳細解説: HTML形式メール作成装飾挿入を中心に、入力取得、処理、出力の順で実行する。
' ファイル利用方法: 対応する処理に必要なシート・セル・選択範囲・ファイルを準備し、マクロを実行する。
' ファイル利用環境: Windows版Excel、クラシックOutlook、Outlookプロファイル。
' ファイル実行例: ExcelのAlt+F8から「HTML形式メール作成装飾挿入」を選択して実行する。

' ── HTML形式メール作成装飾挿入 ──
Public Sub HTML形式メール作成装飾挿入()
    Dim outlook As Object: Set outlook = GetOutlook()
    Dim mail As Object: Set mail = outlook.CreateItem(0)
    mail.To = InputBox("宛先を入力してください", "HTMLメール")
    mail.Subject = InputBox("件名を入力してください", "HTMLメール")
    mail.HTMLBody = "<html><body><h2 style='color:#2F5597'>" & _
                    Replace(mail.Subject, "&", "&amp;") & _
                    "</h2><table border='1' cellpadding='6' style='border-collapse:collapse'>" & _
                    "<tr style='background:#D9EAF7'><th>項目</th><th>内容</th></tr>" & _
                    "<tr><td>本文</td><td>" & Replace(InputBox("本文を入力", "HTMLメール"), vbCrLf, "<br>") & _
                    "</td></tr></table></body></html>"
    mail.Save
End Sub




