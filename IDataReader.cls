VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''Additional References Required:
''Microsoft Scripting Runtime

''Additional Classes Required:
''None

Dim dataReader As IDataReader

Public Sub ReadFile(tStream As Scripting.TextStream): End Sub

Public Function GetRecordset() As Scripting.Dictionary: End Function

Public Property Get FieldNames() As String(): End Property
