Attribute VB_Name = "FunctionCDOEmail"
Option Compare Database
Option Explicit
Option Base 0

Sub testCDOEmail()

Dim EMailTo As String
Dim EMailFrom As String
Dim EMailCC As String
Dim EMailBCC As String
Dim EMailSubject As String
Dim EmailHTMLBody As String
Dim EMailTextBody As String
Dim EMailAttachment As String

    EMailTo = "christopher_palmer@cwb.ca"
    'EMailFrom = "aritchie@p-w-t.ca"
    'EMailFrom = "andrew.nelson@gst.ca"
    'EMailFrom = "melinda.roast@gst.ca"
    'EMailFrom = "jfehr@p-w-t.ca"
    EMailCC = ""
    EMailBCC = ""
    EMailSubject = "quote RSN5"
    EmailHTMLBody = ""
    EMailTextBody = "Original quote request send at " & Now()
    EMailAttachment = ""
    
    If funCDOEmail(EMailTo:=EMailTo, EMailFrom:=EMailFrom, EMailCC:=EMailCC, EMailBCC:=EMailBCC, EMailSubject:=EMailSubject, EmailHTMLBody:=EmailHTMLBody, EMailTextBody:=EMailTextBody, EMailAttachment:=EMailAttachment) = True Then
        MsgBox ("Success!")
    Else
        MsgBox ("Failure!")
    End If

End Sub


Function funCDOEmail(EMailTo As String, EMailFrom As String, Optional EMailCC As String, Optional EMailBCC As String, Optional EMailSubject As String, Optional EmailHTMLBody As String, Optional EMailTextBody As String, Optional EMailAttachment As String)
On Error GoTo errorCDOEmail
'
Dim MessageObject As Object
Dim configurationobject As Object
'
'Instantiate the SMTP COM's Objects.
Set MessageObject = CreateObject("CDO.Message")
Set configurationobject = CreateObject("CDO.Configuration")
    configurationobject.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    configurationobject.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.wpg.cwb.ca"
    configurationobject.Fields.Update
'Create the e-mail
Set MessageObject.Configuration = configurationobject
    MessageObject.To = EMailTo
    MessageObject.FROM = EMailFrom
' the remaining are optional
    If IsMissing(EMailCC) Or EMailCC = "" Then: Else: MessageObject.CC = EMailCC
    If IsMissing(EMailBCC) Or EMailBCC = "" Then: Else: MessageObject.BCC = EMailBCC
    If IsMissing(EMailSubject) Or EMailSubject = "" Then: Else: MessageObject.Subject = EMailSubject
    If IsMissing(EmailHTMLBody) Or EmailHTMLBody = "" Then: Else: MessageObject.HTMLBody = EmailHTMLBody
    If IsMissing(EMailTextBody) Or EMailTextBody = "" Then: Else: MessageObject.TextBody = EMailTextBody
    If IsMissing(EMailAttachment) Or EMailAttachment = "" Then: Else: MessageObject.addattachment (EMailAttachment)
    
    MessageObject.Send
' clean up objects
Set MessageObject = Nothing
Set configurationobject = Nothing
'
funCDOEmail = True
Exit Function
errorCDOEmail:
Set MessageObject = Nothing
Set configurationobject = Nothing

funCDOEmail = False
End Function

