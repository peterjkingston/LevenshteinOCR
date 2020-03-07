VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcrobatPublisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''Microsoft Scripting Runtime
''Acrobat OR Acrobat Type 10.0 Library

''Additional Classes Required:
''None

Dim AcPublisher As New CAcrobatPublisher

Private WithEvents e_window As WinSnipPublish
Attribute e_window.VB_VarHelpID = -1
Private WithEvents e_processor As CAcrobatAutoProcessor
Attribute e_processor.VB_VarHelpID = -1
Private m_app As AcroApp, m_avdoc As AcroAVDoc, m_pddoc As AcroPDDoc, tempAVDoc As AcroAVDoc, tempPDDoc As AcroPDDoc

Public Event Published(FilePath As String)

Private Const BLANK_PDF_TEMPLATE_PATH = "Admin Parts\Blank.pdf"

Public Sub Listen(objListenTo As Object)
    If TypeOf objListenTo Is WinSnipPublish Then: Set e_window = objListenTo
    If TypeOf objListenTo Is CAcrobatAutoProcessor Then: Set e_processor = objListenTo
End Sub

Private Sub Class_Initialize()
    Set m_app = New AcroApp
End Sub

Public Property Get App() As AcroApp
    Set App = m_app
End Property

Private Sub CreateFile(startpage As Integer, pagecount As Integer, fromDoc As AcroPDDoc, saveToDirectory As String, tempDir As String, so As String)
    Dim newAVDoc As New AcroAVDoc, newPDDoc As New AcroPDDoc, success As Boolean, checkpagecount As Integer

    On Error GoTo errHandler

    Set newAVDoc = CreateObject("AcroExch.AVDoc")
    Set newPDDoc = newAVDoc.GetPDDoc
    newPDDoc.Create
    
    newPDDoc.InsertPages -1, fromDoc, startpage - 1, pagecount, 0
    checkpagecount = newPDDoc.GetNumPages
    success = newPDDoc.Save(PDSaveFull, saveToDirectory)
    
    If success Then
        RaiseEvent Published(saveToDirectory)
    End If
    
    WrapUp newPDDoc, newAVDoc
    Exit Sub
    
errHandler:
    WrapUp newPDDoc, newAVDoc
    MsgBox "There was an error creating the file." & vbCrLf & Err.Description, vbOKOnly, "Error " & Err.Number
    Err.clear
End Sub

Private Sub Class_Terminate()
    WrapUp tempPDDoc, tempAVDoc
    WrapUp m_pddoc, m_avdoc
    
    If IsObject(m_app) Then: m_app.Exit
    Set m_app = Nothing
End Sub

Private Sub WrapUp(pddoc As AcroPDDoc, avdoc As AcroAVDoc)
    
    If IsObject(pddoc) And Not pddoc Is Nothing Then: m_pddoc.Close
    If IsObject(avdoc) And Not pddoc Is Nothing Then: m_avdoc.Close -1

    Set pddoc = Nothing
    Set avdoc = Nothing
    
End Sub

Private Sub e_processor_SnipRequested(attach As String, so As String, documentType As String, startPageNum As Integer, endPageNum As Integer, eof As Boolean)
    RequestSnip attach, so, documentType, startPageNum, endPageNum, eof
End Sub

Private Sub e_window_SnipRequested(attach As String, so As String, documentType As String, startPageNum As Integer, endPageNum As Integer, eof As Boolean)
    Dim mainDirectory As String, lengthPages As Integer, fso As Scripting.FileSystemObject
    
    On Error GoTo errHandler
    
    Set m_pddoc = GetPDDocFromFile(attach)
    
    lengthPages = GetEndPages(eof, endPageNum) - startPageNum
    If Not m_pddoc Is Nothing Then
        CreateFile startPageNum, lengthPages, m_pddoc, GetPublishDirectory(so & " - " & documentType, so), mainDirectory, so
    End If
    
    m_pddoc.Close
    Kill mainDirectory
errHandler:
    Set fso = New FileSystemObject
    If Not m_pddoc Is Nothing Then
        m_pddoc.Close
        If fso.FileExists(mainDirectory) Then: Kill mainDirectory
    End If
End Sub

Public Sub RequestSnip(attach As String, so As String, documentType As String, startPageNum As Integer, endPageNum As Integer, eof As Boolean)
    e_window_SnipRequested attach, so, documentType, startPageNum, endPageNum, eof
End Sub

Private Function Validate(avdoc As AcroAVDoc, directory As String, Optional windowTitle As String = "tempPDF") As Boolean
    Dim isOpen As Boolean
    isOpen = avdoc.Open(directory, windowTitle)
    If isOpen Then
        If avdoc.IsValid Then
            Validate = True
        Else
            MsgBox "Not a valid PDF File.", vbOKOnly, ""
            Validate = False
        End If
    Else
        MsgBox "Acrobat failed to open the file."
        Validate = False
    End If
End Function

Private Function GetEndPages(eof As Boolean, pageNum As Integer) As Integer
    If eof Then
        GetEndPages = m_pddoc.GetNumPages + 1
    Else
        GetEndPages = pageNum
    End If
End Function

Private Function GetPDDocFromFile(directory As String) As AcroPDDoc
    Set m_avdoc = CreateObject("AcroExch.AVDoc")
    If m_avdoc.Open(directory, "mainTemp") Then
        Set GetPDDocFromFile = m_avdoc.GetPDDoc
    Else
        Set GetPDDocFromFile = Nothing
    End If
End Function

Private Function GetPublishDirectory(fileLabel As String, so As String) As String
    Dim pubDir As String, fso As FileSystemObject
    
    Set fso = New FileSystemObject
    pubDir = CONSTANTS.FILE_ROOT & "\" & CONSTANTS.GetDepartment(VBA.Environ$("username")) & "\" & so
    
    If Not fso.FolderExists(pubDir) Then: fso.CreateFolder pubDir
    
    GetPublishDirectory = pubDir & "\" & fileLabel & ".pdf"
    
End Function
