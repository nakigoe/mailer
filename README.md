# Microsoft Access VBA E-mailer

VBA mailer for Microsoft Access to send HTML Emails to the mailing list in the database

Utilizes Your Gmail account (set-up a special app password), see appropriate Google instructions for authorizing a custom app to send emails.

## Usage Instructions

- Change `username` to Your Gmail
- Inside Your Goolge account find a setting in the Security Tab > Two Factor Authentication > ` Allow and app access my accont` and let Google generate a single-usage password for this exact app
- Change `password` to the Google generated password for the app
- Change address to Your HTML Email on disk inside `contentFilePath`
- Optional: uncomment a file attachment line and change address to Your attachment inside `.AddAttachment`

```
Option Compare Database
Option Explicit

Sub SendEmails()
    ' Gmail account credentials
    Dim username As String
    Dim password As String
    
    ' Load Gmail account credentials
    username = "mygmail@gmail.com"
    password = "google_generated_password_for_an_app"
    
    ' Gmail SMTP server settings
    Const ServerName As String = "smtp.gmail.com"
    Const ServerPort As Integer = 465
    
    ' Email parameters
    Dim strRecipient As String
    Dim strSubject As String
    Dim strBody As String
    
    ' Load email content from external HTML file
    Dim contentFilePath As String
    contentFilePath = "D:\Docs\test.html"
    
    ' Dim contentFile As Integer
    ' contentFile = FreeFile
    ' Open contentFilePath For Input As #contentFile
    ' strBody = Input$(LOF(contentFile), contentFile)
    ' Close #contentFile
    
    Dim theStream As Object
    Set theStream = CreateObject("ADODB.Stream")
    
    With theStream
        .Charset = "UTF-8"
        .Type = 2 'adTypeText
        .Open
        .LoadFromFile contentFilePath
        strBody = .ReadText
        .Close
    End With
    
    ' Access database table containing contact information
    Dim rsContacts As Recordset
    Set rsContacts = CurrentDb.OpenRecordset("Send HTML Emails to the List")
    
    ' CDO message object
    Dim cdoMessage As Object
    Set cdoMessage = CreateObject("CDO.Message")
    
    ' CDO configuration object
    Dim cdoConfig As Object
    Set cdoConfig = CreateObject("CDO.Configuration")
    
    ' Set configuration properties
    With cdoConfig.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = username
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = ServerName
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = ServerPort
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Update
    End With
    
    ' Loop through the contacts and send emails
    rsContacts.MoveFirst
    Do Until rsContacts.EOF
        strRecipient = rsContacts("Email")
        strSubject = "Hey, How About VBA?"
        
        ' Compose the email
        With cdoMessage
            .BodyPart.Charset = "utf-8"
            .To = strRecipient
            .From = username
            .Subject = strSubject
            .HTMLBody = strBody
            .Configuration = cdoConfig
            ' .AddAttachment "D:\Docs\CV Japanese style.pdf"
            .Send
        End With
        
        rsContacts.MoveNext
    Loop
    
    ' Cleanup
    rsContacts.Close
    Set rsContacts = Nothing
    Set cdoMessage = Nothing
    Set cdoConfig = Nothing
    
    MsgBox "Emails sent successfully!"
End Sub
```

Insert the Macros into the Access database with emails, pick a file with HTML Email, press a button, send it all!!!
