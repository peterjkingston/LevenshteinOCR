VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSAPOrderReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataReader

Dim SAPOrderReader As New CSAPOrderReader
Private m_dict As Scripting.Dictionary, m_fields() As String

Private Sub Class_Initialize()
    Set m_dict = New Scripting.Dictionary
End Sub

Public Sub IDataReader_ReadFile(tStream As Scripting.TextStream)
    Dim line As String, delimiter As String, data() As String, keyColumn As Integer
    
    delimiter = "|"
    keyColumn = 1
    
    Do Until tStream.AtEndOfStream
        line = tStream.ReadLine
        If VBA.left(line, 1) = delimiter Then
            If (Not m_fields) = -1 Then
                m_fields = Split(line, delimiter)
            Else
                data = Split(line, delimiter)
                If Not m_dict.Exists(Trim(data(keyColumn))) Then
                    m_dict.Add Trim(data(keyColumn)), data
                End If
            End If
        End If
    Loop
    
End Sub

Public Function IDataReader_GetRecordset() As Scripting.Dictionary
    
    Dim dict_copy As Scripting.Dictionary, keys As Variant, records As Variant, i As Long
    
    Set dict_copy = New Scripting.Dictionary
    keys = m_dict.keys
    records = m_dict.items
    
    For i = 0 To UBound(keys)
        dict_copy.Add keys(i), records(i)
    Next i
    
    Set IDataReader_GetRecordset = dict_copy
End Function

Public Property Get IDataReader_FieldNames() As String()
    IDataReader_FieldNames = m_fields
End Property
