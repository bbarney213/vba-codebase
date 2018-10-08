VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries.Interfaces")

Option Explicit

Public Sub LoadRecordsFromData(ByRef InputData As Variant, Optional ByRef RecordDriver As IHeaderDriver = Nothing)

End Sub

Public Property Get FilteredRecords(ByRef Queries As IQueryCollection) As IRecordCollection

End Property

Public Property Get Records() As Collection

End Property

Public Property Set Records(value As Collection)

End Property

Public Property Get Data(Optional ByRef ExcludeHeaders As Boolean = False, Optional ByRef OutputHeadersToUse As Scripting.Dictionary = Nothing) As Variant

End Property
