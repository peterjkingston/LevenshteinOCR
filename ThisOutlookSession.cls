VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim cdt() As String, mail As MailItem
    
    cdt = Strings.Split(EntryIDCollection, ",")
    
    Dim index As Integer
    For index = 0 To UBound(cdt)
        Set mail = TryGetMail(cdt(index))
        If Not mail Is Nothing Then
            If TypeOf mail Is MailItem Then
                VBARules.Run mail
            End If
        End If
    Next index
End Sub

Private Function TryGetMail(entryID As String) As MailItem
    On Error Resume Next
    Set TryGetMail = Outlook.Session.GetItemFromID(entryID)
End Function

