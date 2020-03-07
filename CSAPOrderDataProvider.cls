VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSAPOrderDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''Microsoft Scripting Runtime

''Additional Classes Required:
''IDataProvider

Implements IDataProvider

Dim SAPOrderDataProvider As New CSAPOrderDataProvider
Private m_dict As Scripting.Dictionary, m_fields() As String

Public Function IDataProvider_GetData(fieldName As String, key As String, Optional keyField As String = "Document") As String
    Dim data() As String, dataPiece As String, keyfieldIndex As Integer, i As Integer
    
    If keyField = "Document" Then
        If m_dict.Exists(key) Then
            data = m_dict(key)
            If FieldExists(fieldName) Then
                dataPiece = data(IndexOf(fieldName))
                IDataProvider_GetData = Trim(dataPiece)
            End If
        End If
    Else
        If m_dict.Exists(key) Then
            keyfieldIndex = IndexOf(keyField)
            data = m_dict.items
            For i = 0 To UBound(data, 1)
                If (data(keyfieldIndex, i) = key) Then
                    dataPiece = data(IndexOf(fieldName), i)
                    IDataProvider_GetData = Trim(dataPiece)
                    Exit For
                End If
            Next i
        End If
    End If
End Function

Public Function IDataProvider_Exists(fieldName As String, key As String, Optional keyField As String = "Document") As Boolean
    Dim data() As String
    
    If m_dict.Exists(key) Then
        data = m_dict(key)
        If FieldExists(fieldName) Then
            If Trim(data(IndexOf(fieldName))) <> "" Then
                IDataProvider_Exists = True
            End If
        End If
    End If
    
End Function

Private Function FieldExists(fieldName As String) As Boolean
    Dim i As Integer, j As Integer, poHeaders As Collection
    Set poHeaders = New Collection
    
    If fieldName = "PO number" Then
        poHeaders.Add "Purchase order number"
        poHeaders.Add "Purchase order no."
    End If
    
    For i = 0 To UBound(m_fields)
        If poHeaders.count > 0 Then
            For j = 1 To poHeaders.count
                If m_fields(i) = fieldName Or m_fields(i) = poHeaders(j) Then
                    FieldExists = True
                    Exit For
                End If
            Next j
        Else
            If m_fields(i) = fieldName Then
                FieldExists = True
                Exit For
            End If
        End If
    Next i
    
End Function

Private Function IndexOf(fieldName As String) As Integer
    Dim i As Integer, j As Integer, poHeaders As Collection
    Set poHeaders = New Collection
    
    If fieldName = "PO number" Then
        poHeaders.Add "Purchase order number"
        poHeaders.Add "Purchase order no."
    End If
    
    If poHeaders.count > 0 Then
        For i = 0 To UBound(m_fields)
            For j = 1 To poHeaders.count
                If m_fields(i) = fieldName Or _
                   m_fields(i) = poHeaders(j) Then
                    IndexOf = i
                    Exit For
                End If
            Next j
        Next i
    Else
        For i = 0 To UBound(m_fields)
            If m_fields(i) = fieldName Then
                IndexOf = i
                Exit For
            End If
        Next i
    End If
End Function

Public Property Get IDataProvider_FieldNames() As String()
    IDataProvider_FieldNames = m_fields
End Property

Private Sub Class_Initialize()
    UpdateData
End Sub
Private Sub UpdateData()
    Dim dataReader As New IDataReader, fso As Scripting.FileSystemObject, tStream As Scripting.TextStream
    
    Set dataReader = New CSAPOrderReader
    Set fso = New Scripting.FileSystemObject
    
    Set tStream = fso.OpenTextFile(CONSTANTS.LOGPATH)
    dataReader.ReadFile tStream
    tStream.Close
    
    Set tStream = fso.OpenTextFile(CONSTANTS.LOGPATHCA)
    dataReader.ReadFile tStream
    tStream.Close
    
    m_fields = TrimStrings(dataReader.FieldNames)
    Set m_dict = dataReader.GetRecordset
End Sub

Private Function TrimStrings(stringAry() As String) As String()
    Dim i As Integer, newAry() As String
    
    ReDim newAry(UBound(stringAry)) As String
    
    For i = 0 To UBound(newAry)
        newAry(i) = Trim(stringAry(i))
    Next i
    
    TrimStrings = newAry
End Function
