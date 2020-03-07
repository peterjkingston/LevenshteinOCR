VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CShortcutPublisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''Microsoft Scripting Runtime
''Windows Script Host Object Model

''Additional Classes Required:
''WinSnipPublish
''WinCSPublish

Dim shortcutPublisher As New CShortcutPublisher

Private WithEvents e_winSnip As WinSnipPublish
Attribute e_winSnip.VB_VarHelpID = -1
Private WithEvents e_winCS  As WinCSPublish
Attribute e_winCS.VB_VarHelpID = -1
Private WithEvents e_filePublisher As CFilePublisher
Attribute e_filePublisher.VB_VarHelpID = -1

Public Sub Listen(objListenTo As Object)
    If TypeOf objListenTo Is WinSnipPublish Then: Set e_winSnip = objListenTo
    If TypeOf objListenTo Is WinCSPublish Then: Set e_winCS = objListenTo
    If TypeOf objListenTo Is CFilePublisher Then: Set e_filePublisher = objListenTo
End Sub

Private Sub e_filePublisher_Published(FilePath As String)
    PublishShortcut FilePath
End Sub

Private Sub e_winCS_PublishRequested(FilePath As String, label As String)
    PublishShortcut FilePath & "\" & label
End Sub

Private Sub e_winSnip_SnipRequested(attach As String, so As String, documentType As String, startPageNum As Integer, endPageNum As Integer, eof As Boolean)
    PublishShortcut CONSTANTS.FILE_ROOT & "\Administration\" & so & "\" & so & " - " & documentType & ".pdf"
End Sub

Public Sub PublishShortcut(FilePath As String)
    Dim newFilePath As String, fso As FileSystemObject, filename As String, oFile As file, salesOrder As String
    Set fso = New FileSystemObject
    Set oFile = fso.GetFile(FilePath)
    filename = GetShorterName(oFile.Name)
    salesOrder = Split(filename, " ")(0)
    newFilePath = CONSTANTS.FILE_ROOT & "\Sales Public\" & salesOrder & "\" & filename
    
    SubmitPUTShortcut FilePath, newFilePath
End Sub
Public Function SubmitPUTShortcut(oldFilePath As String, newFilePath As String) As Boolean
    Dim wShell As IWshRuntimeLibrary.WshShell, shortcut As IWshRuntimeLibrary.WshShortcut
    Set wShell = New WshShell
    Set shortcut = wShell.CreateShortcut(newFilePath)
    
    With shortcut
        .TargetPath = oldFilePath
        .WorkingDirectory = oldFilePath
        If TryCreateFolder(newFilePath) Then
            .Save
        End If
    End With
    
    SubmitPUTShortcut = True
End Function
Private Function TryCreateFolder(newFilePath As String) As Boolean
    Dim fso As FileSystemObject, folderpath As String
    Set fso = New FileSystemObject
    
    folderpath = fso.GetParentFolderName(newFilePath)
    If Not fso.FolderExists(folderpath) Then
        fso.CreateFolder folderpath
    End If
    
    TryCreateFolder = True
End Function
Private Function GetShorterName(filename As String) As String
    Dim fso As Scripting.FileSystemObject, extension As String
    
    Set fso = New Scripting.FileSystemObject
    extension = fso.GetExtensionName(filename)
    GetShorterName = VBA.left(filename, Len(filename) - Len(extension)) & "lnk"
    
End Function
