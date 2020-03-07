VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCOADistributor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''Microsoft Outlook 16.0 Object Library
''Microsoft Scripting Runtime

''Additional Classes Required:
''WinSnipPublish
''CONSTANTS

Dim COADistributor As New CCOADistributor

Private WithEvents e_winSnip As WinSnipPublish
Attribute e_winSnip.VB_VarHelpID = -1

Private m_dataProvider As IDataProvider

Public Sub Listen(objListenTo As Object)
    If TypeOf objListenTo Is WinSnipPublish Then: Set e_winSnip = objListenTo
End Sub

Public Sub SetDataProvider(objDataProvider As IDataProvider)
    Set m_dataProvider = objDataProvider
End Sub

Private Sub m_winSnip_SnippingComplete(snip As IPDFSnip, args As Variant)
    Dim mail As MailItem, soldTo As String, salesOrder As String
    With m_dataProvider
        salesOrder = snip.GetSO
        
        If .Exists("EmailDistro", salesOrder) Then
            Set mail = Outlook.CreateItem(olMailItem)
            mail.To = .GetData("EmailDistro", salesOrder)
            mail.Subject = "PO# " & .GetData("PO number", salesOrder)
            mail.Attachments.Add CONSTANTS.FILE_ROOT & "\Administration\" & salesOrder & "\" & salesOrder & " - Certificate of Analysis.pdf"
            
            mail.HTMLBody = GetHTMLBody
            mail.Display
        End If
    End With
End Sub

Private Function GetHTMLBody() As String
    Dim fso As Scripting.FileSystemObject, tStream As Scripting.TextStream
    
    Set fso = New Scripting.FileSystemObject
    Set tStream = fso.OpenTextFile(CONSTANTS.FILE_ROOT & "\Admin Parts RO\COADistro\HTMLMessage.txt")
    
    GetHTMLBody = tStream.ReadAll
    tStream.Close
End Function
