Sub generateMails()

Dim objOutlook As Object
Dim objMail As Object

Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)

mailBetreff = ActiveSheet.Cells(5, 2).Value
mailBody = ActiveSheet.Cells(5, 3).Value

mailCC1 = ActiveSheet.Cells(7, 3).Value
mailCC2 = ActiveSheet.Cells(8, 3).Value
mailCC3 = ActiveSheet.Cells(9, 3).Value
mailCC = mailCC1 & ";" & mailCC2 & ";" & mailCC3

mailBCC1 = ActiveSheet.Cells(13, 3).Value
mailBCC2 = ActiveSheet.Cells(14, 3).Value
mailBCC3 = ActiveSheet.Cells(15, 3).Value
mailBCC4 = ActiveSheet.Cells(16, 3).Value
mailBCC5 = ActiveSheet.Cells(17, 3).Value
mailBCC6 = ActiveSheet.Cells(18, 3).Value
mailBCC = mailBCC1 & ";" & mailBCC2 & ";" & mailBCC3 & ";" & mailBCC4 & ";" & mailBCC5 & ";" & mailCC6

iRow = 8

Do While Not IsEmpty(ActiveSheet.Cells(iRow, 6))

    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    mailSubject = ActiveSheet.Cells(iRow, 6)
    
    mailAddress1 = ActiveSheet.Cells(iRow, 7)
    mailAddress2 = ActiveSheet.Cells(iRow, 8)
    mailAddress3 = ActiveSheet.Cells(iRow, 9)
    mailAddress = mailAddress1 & ";" & mailAddress2 & ";" & mailAddress3
    
    mailText = mailSubject & vbLf & vbLf & mailBody
    
    With objMail
        .To = mailAddress
        .CC = mailCC
        .BCC = mailBCC
        .Subject = mailBetreff
        .Body = mailText
        If Not IsEmpty(ActiveSheet.Cells(iRow, 10).Value) Then
                .Attachments.Add ActiveSheet.Cells(iRow, 10).Value
        End If
        .Display
                
    End With
    
    iRow = iRow + 1
Loop


End Sub

Sub addAttachment()

Dim exl As Object
Set exl = CreateObject("Excel.Application")

ExcelFile = exl.Application.GetOpenFilename(, , "Anhang auswählen")
ActiveSheet.Cells(ActiveCell.Row, 10).Value = ExcelFile

End Sub
