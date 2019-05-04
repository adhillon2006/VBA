Sub Email_Attachment
Dim OutlookApp As Object
Dim OutlookMail As Object

'Save workbook before sending
ActiveWorkbook.Save

'Creating objects to control outlook
Set OutlookApp = CreateObject("Outlook.Application")
Set OutlookMail = OutlookApp.CreateItem(0)

On Error Resume Next

With OutlookMail
    
    .To = "Insert email"
    .CC = ""
    .BCC = ""
    .Subject = "Insert subject"
    .HTMLBody = "<body><p>You can find the required information at " & _
                              "<a href='http://infotest.com/12345/12345.html'>" & _
                              "http://infotest.com/12345/12345.html</a>. Good luck!"
    .Send

End With

Set OutlookMail = Nothing
Set OutlookApp = Nothing

End Sub
