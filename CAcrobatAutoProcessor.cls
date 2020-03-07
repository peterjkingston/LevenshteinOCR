VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcrobatAutoProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim processor As CAcrobatAutoProcessor

Private Enum LISTCOLUMNS
    KEY_TYPE = 0
    KEY_NUM
    DOC_TYPE
    START_PAGE
End Enum
Private Enum KEY_TYPES
    SALES_ORDER = 0
    PURCHASE_ORDER
    DELIVERY_NUM
End Enum

Private m_Activated As Boolean, m_mail As MailItem, m_dataProvider As IDataProvider

Public Event SnipRequested(attach As String, so As String, documentType As String, startPageNum As Integer, endPageNum As Integer, eof As Boolean)
Public Event SnipIterated(precentComplete As Integer, message As String)
Public Event SnippingComplete(snip As CAcrobatAutoProcessor, args As Variant)
Public Event SnipAfter(so As String)
Public Event BeforeSnipping(snip As CAcrobatAutoProcessor, args As Variant)

Private m_app As AcroApp, m_publisher As CAcrobatPublisher, m_SO As String
Private m_dataOrderProvider As CSAPOrderDataProvider, m_logger As CSalesOrderLogger, m_mailchanger As CMailInfoChanger, m_mailitem As MailItem, m_postOCR As CAcrobatPostOCR, m_pddoc As AcroPDDoc, m_avdoc As AcroAVDoc

Public Function Run(mail As MailItem) As Boolean
    Dim docPageTypes As Scripting.Dictionary, snip As Integer, nextpagestart As Integer, doctype As String, docpage As Integer
    
    On Error GoTo ErrorHandler
    Set m_mailitem = mail
    Set docPageTypes = GetDocPagesTypes
    
    For snip = 0 To docPageTypes.count - 1
        RaiseEvent BeforeSnipping(Me, 0)
        If snip = docPageTypes.count - 1 Then
            nextpagestart = -1
        Else
            nextpagestart = docPageTypes.keys(snip + 1)
        End If
        doctype = docPageTypes.items(snip)
        docpage = docPageTypes.keys(snip)
        RaiseEvent SnipRequested(CONSTANTS.FILE_ROOT & "\tempOCR.pdf", so, doctype, docpage, nextpagestart, snip = docPageTypes.count - 1)
        RaiseEvent SnipAfter(so)
        RaiseEvent SnipIterated((snip / docPageTypes.count) * 100, "")
    Next snip
    RaiseEvent SnippingComplete(Me, 0)
    
    Run = True
    Exit Function
ErrorHandler:
    Run = False
    Exit Function
End Function

Public Property Get mail() As MailItem
    Set mail = m_mailitem
End Property

Public Property Get so() As String
    If m_SO = "" Then: m_SO = left(m_postOCR.GetSO(), 7)
    so = m_SO
End Property

Public Function GetDocPagesTypes() As Scripting.Dictionary
    Dim page As Integer, pagecount As Integer, doctype As String, skipthis As Boolean, resultDict As Scripting.Dictionary, avdoc As AcroAVDoc, lastdocType As String
    
    Set resultDict = New Scripting.Dictionary
    Set m_pddoc = GetPDDocFromFile(CONSTANTS.FILE_ROOT & "\tempOCR.pdf")
    
    pagecount = m_pddoc.GetNumPages()
    m_postOCR.Setup CONSTANTS.FILE_ROOT & "\tempOCR.pdf"
    
    For page = 0 To pagecount - 1
        skipthis = False
        lastdocType = doctype
        Select Case m_postOCR.GetPageID(page)
            Case 0 ''Delivery Note
                doctype = "Delivery Note"
            Case 1 ''Bill of Lading
                doctype = "Bill of Lading"
            Case 2 ''Certificate of Analysis
                doctype = "Certificate of Analysis"
            Case 3 ''SealManifest
                doctype = "Seal Manifest"
            Case Else
                skipthis = True
        End Select
        If doctype = lastdocType Then: skipthis = True
        If Not skipthis Then
            resultDict.Add page + 1, doctype
        End If
    Next page
    
    
    Set GetDocPagesTypes = resultDict
End Function

Private Function GetPDDocFromFile(directory As String) As AcroPDDoc
    Set m_avdoc = CreateObject("AcroExch.AVDoc")
    If m_avdoc.Open(directory, "autoTemp") Then
        Set GetPDDocFromFile = m_avdoc.GetPDDoc
    Else
        Set GetPDDocFromFile = Nothing
    End If
End Function

Private Sub Class_Initialize()
    Set m_dataOrderProvider = New CSAPOrderDataProvider
    
    Set m_publisher = New CAcrobatPublisher
    m_publisher.Listen Me
    
    Set m_app = New AcroApp
    
    Set m_logger = New CSalesOrderLogger
    m_logger.Listen Me
    
    Set m_mailchanger = New CMailInfoChanger
    m_mailchanger.Listen Me
    m_mailchanger.SetDataProvider m_dataOrderProvider
    
    Set m_postOCR = New CAcrobatPostOCR
    m_postOCR.SetDataProvider m_dataOrderProvider
End Sub

Public Sub CloseObject()
    Class_Terminate
End Sub

Private Sub Class_Terminate()
    If Not m_app Is Nothing Then
        m_pddoc.Close
        m_avdoc.Close -1
        m_app.Hide
        m_app.CloseAllDocs
        Set m_app = Nothing
    End If
    Set m_publisher = Nothing
    Set m_postOCR = Nothing
End Sub
