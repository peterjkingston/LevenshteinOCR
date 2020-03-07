VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcrobatPostOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum PageID
    csUNKDocument = -1
    csDeliveryNote = 0
    csInternalBOL
    csCertificateAnalysis
    csSealManifest
End Enum

Const delvNote As String = "DELIVERY NOTE"
Const intBOL As String = "BILL OF LADING"
Const certAnal As String = "Date: ??/??/??"

Const sealMan As String = "DRUM SEAL MANIFEST"

Dim postOCR As CAcrobatPostOCR
Private m_pddoc As AcroPDDoc, m_app As AcroApp, m_dataProvider As IDataProvider, m_avdoc As AcroAVDoc
Public Property Set App(App As AcroApp)
    Set m_app = New AcroApp
End Property

Public Sub SetDataProvider(objDataProvider As Object)
    If TypeOf objDataProvider Is IDataProvider Then: Set m_dataProvider = objDataProvider
End Sub

Public Sub SetPDDoc(pddoc As AcroPDDoc)
    Set m_pddoc = pddoc
End Sub

Public Sub Setup(path As String)
    Set m_pddoc = GetPDDocFromFile(path)
End Sub

Private Function GetPDDocFromFile(directory As String) As AcroPDDoc
    Set m_avdoc = CreateObject("AcroExch.AVDoc")
    If m_avdoc.Open(directory, "ocrTemp") Then
        Set GetPDDocFromFile = m_avdoc.GetPDDoc
    Else
        Set GetPDDocFromFile = Nothing
    End If
End Function

Public Function GetSO() As String
    Dim page As Integer, so As String, id As PageID, pddoc As AcroPDDoc, regex As RegExp
    
    Set pddoc = m_pddoc
    Set regex = New RegExp
    regex.Pattern = "\s*\d{7}\s*"

    For page = 0 To pddoc.GetNumPages - 1
        id = GetPageID(page)
        Select Case id
            Case csDeliveryNote
                so = RemoveLeadingZeros(SearchAroundRegex(680, 669, 124, 173, page, regex))
                If (Len(Trim(so)) <> 7) Then: so = RemoveLeadingZeros(SearchAroundRegex(700, 680, 124, 173, page, regex))
            Case csInternalBOL
                so = RemoveLeadingZeros(SearchAroundRegex(578, 565, 463, 504, page, regex))
            Case csCertificateAnalysis
                so = RemoveLeadingZeros(SearchAroundRegex(538, 526, 421, 457, page, regex))
        End Select
        If ValidateSO(so, pddoc, id) Then: Exit For
    Next page
    
    GetSO = so
End Function

Public Function ValidateSO(so As String, pddoc As AcroPDDoc, id As PageID) As Boolean
    
    Dim result As Boolean: result = False
    Dim isSO As RegExp: Set isSO = New RegExp: isSO.Pattern = "\d{" & CONSTANTS.SO_LENGTH & "}"
    Dim isTooLong As RegExp: Set isTooLong = New RegExp: isTooLong.Pattern = "\d{" & CONSTANTS.SO_LENGTH + 1 & "}"
    Dim notNumbers As RegExp: Set notNumbers = New RegExp: notNumbers.Pattern = "[a-z,0,\n]*"
    Dim replaceClear As String: replaceClear = ""
    Dim page As Integer: page = 0
    
    so = notNumbers.Replace(so, replaceClear)
    If isSO.Test(so) And Not isTooLong.Test(so) Then
        so = left(so, CONSTANTS.SO_LENGTH)
        If m_dataProvider.Exists("Name 1", so) Then
                Dim dataName1 As String: dataName1 = m_dataProvider.GetData("Name 1", so)
                For page = 0 To pddoc.GetNumPages - 1
                    Dim readName1 As String: readName1 = GetName1(GetPageID(page), page, so)
                    result = isCloseTo(readName1, dataName1, 0.5)
                    If result Then: Exit For
                Next page
        End If
    End If
    
    ValidateSO = result
End Function

Private Function GetName1(id As PageID, page As Integer, so As String) As String
    Dim result As String, expected As String
    
    
    If m_dataProvider.Exists("Document", so) Then
        expected = m_dataProvider.GetData("Name 1", so)
        
        Select Case id
            Case PageID.csDeliveryNote
                result = SearchAround(597, 585, 386, 600, page, expected, 0.5)
                If left(result, 4) = "To: " Then: result = m_dataProvider.GetData("Name 1", Mid(result, 5, 6), "Sold-to pt")
            Case PageID.csInternalBOL
                result = SearchAround(671, 660, 394, 600, page, expected, 0.5)
            Case PageID.csCertificateAnalysis
                result = SearchAround(654, 640, 33, 200, page, expected, 0.5)
        End Select
        
    End If
    
    GetName1 = result
End Function

Function RemoveLeadingZeros(Str As String) As String
  Do While left(Str, 1) = "0" Or left(Str, 1) = ":"
    Str = Mid(Str, 2)
  Loop
  RemoveLeadingZeros = Str
End Function
Public Function GetPageID(page As Integer) As PageID
    Dim id As PageID, i As Integer, threshold As Double
    
    threshold = 0.5
    For i = 0 To 0
        Select Case True
            Case isCloseTo(SearchAround(715, 690, 262, 378, page, delvNote, threshold), delvNote, threshold)
                id = PageID.csDeliveryNote
                Exit For
            Case isCloseTo(SearchAround(1000, 720, 400, 500, page, intBOL, threshold), intBOL, threshold)
                id = PageID.csInternalBOL
                Exit For
            Case isCloseTo(SearchAround(1000, 625, 450, 460, page, certAnal, 0.28), certAnal, 0.28)
                id = PageID.csCertificateAnalysis
                Exit For
            Case isCloseTo(SearchAround(1000, 719, 286, 365, page, sealMan, threshold), sealMan, threshold)
                id = PageID.csSealManifest
                Exit For
            Case Else
                id = csUNKDocument
                Exit For
        End Select
    Next i
    
    GetPageID = id
End Function

Private Function SearchAround(top As Integer, bottom As Integer, left As Integer, right As Integer, page As Integer, expected As String, threshold As Double) As String
    Dim instruction As Integer, found As String, pdrect As AcroRect
    
    Set pdrect = CreateObject("AcroExch.Rect")
    pdrect.top = top
    pdrect.bottom = bottom
    pdrect.left = left
    pdrect.right = right
    
    For instruction = 0 To 4
        Select Case instruction
            Case 0
                found = ReadRect(pdrect, page)
            Case 1
                ScootRect pdrect, 5, 0
                found = ReadRect(pdrect, page)
            Case 2
                ScootRect pdrect, -10, 0
                found = ReadRect(pdrect, page)
            Case 3
                ScootRect pdrect, 5, 5
                found = ReadRect(pdrect, page)
            Case 4
                ScootRect pdrect, 0, -10
                found = ReadRect(pdrect, page)
            Case Else
                found = ""
        End Select
        
        found = TrimAtLineBreak(found)
        If isCloseTo(found, expected, threshold) Then: Exit For
    Next instruction
    
    SearchAround = found
End Function

Private Function TrimAtLineBreak(thisStr As String) As String
    Dim actualLength As Integer, nextLeg As Integer
    
    ''vbcr
    actualLength = InStr(1, thisStr, vbCrLf)
    
    ''vblf
    nextLeg = InStr(1, thisStr, vbLf)
    If actualLength = 0 Then: actualLength = nextLeg
    If nextLeg < actualLength And nextLeg <> 0 Then: actualLength = nextLeg
    
    ''vbcrlf
    nextLeg = InStr(1, thisStr, vbCrLf)
    If nextLeg < actualLength Then: actualLength = nextLeg
    If actualLength = 0 Then: actualLength = nextLeg
    If nextLeg < actualLength And nextLeg <> 0 Then: actualLength = nextLeg
    
    If actualLength > 0 Then
        TrimAtLineBreak = left(thisStr, actualLength)
    Else
        TrimAtLineBreak = thisStr
    End If
End Function

Private Function SearchAroundRegex(top As Integer, bottom As Integer, left As Integer, right As Integer, page As Integer, regex As RegExp) As String
    Dim instruction As Integer, found As String, pdrect As AcroRect
    
    Set pdrect = CreateObject("AcroExch.Rect")
    pdrect.top = top
    pdrect.bottom = bottom
    pdrect.left = left
    pdrect.right = right
    
    For instruction = 0 To 4
        Select Case instruction
            Case 0
                found = ReadRect(pdrect, page)
            Case 1
                ScootRect pdrect, 5, 0
                found = ReadRect(pdrect, page)
            Case 2
                ScootRect pdrect, -10, 0
                found = ReadRect(pdrect, page)
            Case 3
                ScootRect pdrect, 5, 5
                found = ReadRect(pdrect, page)
            Case 4
                ScootRect pdrect, 0, -10
                found = ReadRect(pdrect, page)
            Case Else
                found = ""
        End Select
        
        If regex.Test(found) Then: Exit For
        
    Next instruction
    
    SearchAroundRegex = found
End Function

Public Function ReadRect(pdrect As AcroRect, page As Integer) As String
    
    Dim textSelect As AcroPDTextSelect
    Set textSelect = m_pddoc.CreateTextSelect(page, pdrect)
    If textSelect Is Nothing Then
        ReadRect = ""
    Else
        Dim vector As Integer, Text As String
        For vector = 0 To textSelect.GetNumText() - 1
            Text = Text & textSelect.GetText(vector)
        Next vector
        ReadRect = Text
    End If
End Function

Public Function isCloseTo(thisStr As String, closeTo As String, thresholdpercent As Double) As Boolean
    Dim percent As Double, largestLength As Integer, editDistance As Integer
    
    largestLength = GetLargestLength(thisStr, closeTo)
    editDistance = GetEditDistance(thisStr, closeTo)
    percent = ((largestLength - editDistance) / largestLength)
    
    isCloseTo = percent >= thresholdpercent
End Function
Public Function GetEditDistance(s1 As String, s2 As String)

    Dim i As Integer
    Dim j As Integer
    Dim l1 As Integer
    Dim l2 As Integer
    Dim d() As Integer
    Dim min1 As Integer
    Dim min2 As Integer

    l1 = Len(s1)
    l2 = Len(s2)
    ReDim d(l1, l2)
    For i = 0 To l1
        d(i, 0) = i
    Next
    For j = 0 To l2
        d(0, j) = j
    Next
    For i = 1 To l1
        For j = 1 To l2
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                d(i, j) = d(i - 1, j - 1)
            Else
                min1 = d(i - 1, j) + 1
                min2 = d(i, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                min2 = d(i - 1, j - 1) + 1
                If min2 < min1 Then
                    min1 = min2
                End If
                d(i, j) = min1
            End If
        Next
    Next
    GetEditDistance = d(l1, l2)
End Function
Private Function GetLargestLength(str1 As String, str2 As String) As Integer
    Dim count As Integer
    count = Len(str1)
    If count < Len(str2) Then: count = Len(str2)
    GetLargestLength = count
End Function
Private Sub Class_Terminate()
    If Not m_app Is Nothing Then
        m_pddoc.Close
        m_avdoc.Close -1
        m_app.Hide
        m_app.CloseAllDocs
        Set m_app = Nothing
    End If
End Sub

Private Sub ScootRect(rect As AcroRect, verticalShift As Integer, horizontalShift As Integer)
    rect.top = rect.top + verticalShift
    rect.bottom = rect.bottom + verticalShift
    rect.left = rect.left + horizontalShift
    rect.right = rect.right + horizontalShift
End Sub
