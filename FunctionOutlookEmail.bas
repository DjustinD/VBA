Attribute VB_Name = "FunctionOutlookEmail"
Option Compare Database
Option Explicit

Sub testOutlookEmail()
Dim strEMailTo As String
Dim strEMailCC As String
Dim strEMailBCC As String
Dim strEMailFrom As String
Dim strEMailSubject As String
Dim strEMailBody As String
Dim strEmailAttachmentFiles As String


strEMailTo = "christopher_palmer@cwb.ca"
strEMailCC = "rmg@cwb.ca"
strEMailBCC = "justin_daniels@cwb.ca"
strEMailFrom = "CWB_Market_Research@cwb.ca"
strEMailSubject = "Testing VBA code to mail with Outlook"
strEMailBody = "Testing VBA code to mail with Outlook"
strEmailAttachmentFiles = Array("I:\MktRskRsch\Restricted\MktRsch\SubscriptionDatabase\WaterMarks\justin daniels.pdf", "I:\MktRskRsch\Restricted\MktRsch\SubscriptionDatabase\WaterMarks\Chris Palmer CWB.pdf")
'strEmailAttachmentFiles = "I:\MktRskRsch\Restricted\MktRsch\SubscriptionDatabase\WaterMarks\justin daniels.pdf"
'Call OutlookEmail(EmailTo, EmailCC, EmailBCC, EmailSentOnBehalfOfName, EmailSubject, EmailBody, EmailAttachmentFiles)
Call funOutlookEmail(EMailTo:=strEMailTo, EMailCC:=strEMailCC, EMailBCC:=strEMailBCC, EMailFrom:=strEMailFrom, EMailSubject:=strEMailSubject, EMailBody:=strEMailBody)

End Sub

Function funOutlookEmail(Optional EMailTo As String, Optional EMailCC As String, Optional EMailBCC As String, Optional EMailFrom As String, Optional EMailSubject As String, Optional EMailBody As String, Optional EmailAttachmentFiles, Optional HTMLFlag As Boolean) As Boolean
On Error GoTo ErrorOutlookEmail
'requires reference to "Microsoft Outlook 14.0 Object Library"

Dim MSOutlook As Outlook.Application
Set MSOutlook = CreateObject("Outlook.application")
Dim MSOutlookItem As Outlook.MailItem
Set MSOutlookItem = MSOutlook.CreateItem(olMailItem)
If Not (IsNull(EMailTo) Or EMailTo = "") Then: MSOutlookItem.To = EMailTo
If Not (IsNull(EMailCC) Or EMailCC = "") Then: MSOutlookItem.CC = EMailCC
If Not (IsNull(EMailBCC) Or EMailBCC = "") Then: MSOutlookItem.BCC = EMailBCC
If Not (IsNull(EMailFrom) Or EMailFrom = "") Then: MSOutlookItem.SentOnBehalfOfName = EMailFrom
If Not (IsNull(EMailSubject) Or EMailSubject = "") Then: MSOutlookItem.Subject = EMailSubject
If Not (IsNull(EMailBody) Or EMailBody = "") Then
    If HTMLFlag = True Then
        With MSOutlookItem
            .BodyFormat = olFormatHTML
            .HTMLBody = EMailBody
            .display
        End With
    Else
        MSOutlookItem.body = EMailBody
    End If
End If

If Not IsMissing(EmailAttachmentFiles) Then
    If Not IsNull(EmailAttachmentFiles) Then
        
        Dim MSAttachments As Outlook.Attachments
        Dim ArrayCounter As Integer
        Set MSAttachments = MSOutlookItem.Attachments
        If IsArray(EmailAttachmentFiles) Then
            For ArrayCounter = LBound(EmailAttachmentFiles) To UBound(EmailAttachmentFiles)
                MSAttachments.Add EmailAttachmentFiles(ArrayCounter)
                DoEvents
            Next ArrayCounter
        Else
            MSAttachments.Add EmailAttachmentFiles
        End If
        Set MSAttachments = Nothing
    End If
End If
DoEvents
MSOutlookItem.Send

funOutlookEmail = True
Set MSOutlookItem = Nothing
Set MSOutlook = Nothing

Exit Function

ErrorOutlookEmail:
funOutlookEmail = False
Err.Clear
Set MSOutlookItem = Nothing
Set MSOutlook = Nothing

End Function



