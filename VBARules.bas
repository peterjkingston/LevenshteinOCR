Attribute VB_Name = "VBARules"
Option Explicit

Public Sub Run(mail As MailItem)
    Select Case True
        Case mail.Sender.Address = "theCopier@myCompany.com"
            If Not TryProcessAsOCR(mail) Then: MarkForReview mail
            MoveToFolder mail, "Copier"
            Exit Sub
        Case (mail.Sender.Name = "Peter Kingston" And InStr(1, mail.Body, "theCopier@myCompany.com") > 0)
            If Not TryProcessAsOCR(mail) Then: MarkForReview mail
            Kill CONSTANTS.FILE_ROOT & "\tempOCR.pdf"
            MoveToFolder mail, "Copier"
            Exit Sub
    End Select
End Sub

Public Function TryProcessAsOCR(mail As MailItem) As Boolean
    Dim scans As Attachment
    If mail.Attachments.count > 0 Then
        Set scans = mail.Attachments(1)
        scans.SaveAsFile CONSTANTS.FILE_ROOT & "\tempOCR.pdf"
        
        Dim autoProcessor As CAcrobatAutoProcessor
        Set autoProcessor = New CAcrobatAutoProcessor
        TryProcessAsOCR = autoProcessor.Run(mail)
        autoProcessor.CloseObject
        Set autoProcessor = Nothing
        
    End If
End Function

Public Sub MoveToFolder(mail As MailItem, folderName As String)
    Dim fldr As Outlook.Folder
     Set fldr = Outlook.Session.GetDefaultFolder(olFolderInbox).Parent.Folders(folderName) ''.Folders(folderName)
    mail.Move fldr
End Sub

Public Sub MarkForReview(mail As MailItem)
    mail.Subject = "Flagged for Review"
    mail.Save
End Sub
