VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryableRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries.Interfaces")

Option Explicit

Public Sub AddRecord(ByVal Record As IQueryableRecord)

End Sub

Public Property Get FilteredRecords(ByVal Queries As IQueryCollection) As Collection

End Property

Public Property Get Records() As Collection

End Property

Public Property Set Records(value As Collection)

End Property
