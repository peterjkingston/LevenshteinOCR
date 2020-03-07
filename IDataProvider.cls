VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''None

''Additional Classes Required:
''None

Dim dataProvider As New IDataProvider

Public Function GetData(fieldName As String, key As String, Optional keyField As String = "Document") As String: End Function

Public Function Exists(fieldName As String, key As String, Optional keyField As String = "Document") As Boolean: End Function

Public Property Get FieldNames() As String(): End Property
