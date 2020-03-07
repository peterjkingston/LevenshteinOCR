VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSalesOrderLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''Microsoft Scripting Runtime

''Additional Classes Required:
''WinSnipPublish

Dim SalesOrderLogger As New CSalesOrderLogger

Private WithEvents e_winSnip As WinSnipPublish
Attribute e_winSnip.VB_VarHelpID = -1
Private WithEvents e_processor As CAcrobatAutoProcessor
Attribute e_processor.VB_VarHelpID = -1

Public Sub Listen(objListenTo As Object)
    If TypeOf objListenTo Is WinSnipPublish Then: Set e_winSnip = objListenTo
    If TypeOf objListenTo Is CAcrobatAutoProcessor Then: Set e_processor = objListenTo
End Sub

Private Sub e_winSnip_SnippingComplete(snip As WinSnipPublish, args As Variant)
    Dim salesOrder As String, fileDate As Date, soFilePath As String, fso As Scripting.FileSystemObject, soFileStream As Scripting.TextStream
    Set fso = New FileSystemObject
    soFilePath = "C:\Users\" & VBA.Environ$("Username") & "\Documents\Today's Outlook SOs.txt"
    
    Set soFileStream = fso.OpenTextFile(soFilePath, ForReading, True)
    If Not soFileStream.AtEndOfStream Then: fileDate = CDate(Trim(soFileStream.ReadLine))
    soFileStream.Close
    
    If Not fileDate = Date Then
        ''Keep the file as only the current date.
        Set soFileStream = fso.OpenTextFile(soFilePath, ForWriting)
        soFileStream.WriteLine Date ''<--- Write the date as a header
        soFileStream.Close
    End If
        
    Set soFileStream = fso.OpenTextFile(soFilePath, ForAppending)
    salesOrder = Trim(snip.IPDFSnip_GetSO)
    soFileStream.WriteLine (salesOrder)
    soFileStream.Close
End Sub

Private Sub e_processor_SnippingComplete(snip As CAcrobatAutoProcessor, args As Variant)
    Dim salesOrder As String, fileDate As Date, soFilePath As String, fso As Scripting.FileSystemObject, soFileStream As Scripting.TextStream
    Set fso = New FileSystemObject
    soFilePath = "C:\Users\" & VBA.Environ$("Username") & "\Documents\Today's Outlook SOs.txt"
    
    Set soFileStream = fso.OpenTextFile(soFilePath, ForReading, True)
    If Not soFileStream.AtEndOfStream Then: fileDate = CDate(Trim(soFileStream.ReadLine))
    soFileStream.Close
    
    If Not fileDate = Date Then
        ''Keep the file as only the current date.
        Set soFileStream = fso.OpenTextFile(soFilePath, ForWriting)
        soFileStream.WriteLine Date ''<--- Write the date as a header
        soFileStream.Close
    End If
        
    Set soFileStream = fso.OpenTextFile(soFilePath, ForAppending)
    salesOrder = Trim(snip.so)
    soFileStream.WriteLine (salesOrder)
    soFileStream.Close
End Sub
