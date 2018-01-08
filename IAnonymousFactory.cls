VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnonymousFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.2

'@Folder("CodeBase.Classes.Anonymous.Interfaces")

Option Explicit

Public Sub AddGlobalProperty(ByRef PropertyName As String, ByRef PropertyValue As Variant)

End Sub

Public Sub LoadRecordsFromData(ByRef InputData As Variant, Optional ByRef RecordDriver As IHeaderDriver = Nothing)

End Sub

Public Property Get Records() As Collection

End Property

Public Property Set Records(value As Collection)

End Property

Public Property Get Headers() As Scripting.Dictionary

End Property

Public Property Get Data(Optional ByRef ExcludeHeaders As Boolean = False, Optional ByRef OutputHeadersToUse As Scripting.Dictionary) As Variant

End Property
