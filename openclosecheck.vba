Option Explicit
 
Sub SendMail()

     Dim OApp As Object, OMail As Object, signature As String
Set OApp = CreateObject("Outlook.Application")

Dim sDate As String
    sDate = Format(Now(), "[$-en-US]mmmm d")
    
    Dim sBody As String
    sBody = "Hello," & vbNewLine & vbNewLine & "Please find the opening/closing checklist attached for " & sDate & vbNewLine & vbNewLine & "Have a great day,"
    
Set OMail = OApp.CreateItem(0)
    With OMail
    .Display
    End With
        signature = OMail.Body
    With OMail
    .Subject = "Opening/Closing Checklist " & sDate
         'Specify who it should be sent to
         'Repeat this line to add further recipients
        .Recipients.Add "building.mgmt@mcgill.ca"
         'specify the file to attach
         'repeat this line to add further attachments
       '.Attachments.Add filepath
         'specify the text to appear in the email
        .Body = sBody & signature
         'Choose which of the following 2 lines to have commented out
        .Display
    '.Send
    End With
    
Set OMail = Nothing
Set OApp = Nothing
     
     End Sub
