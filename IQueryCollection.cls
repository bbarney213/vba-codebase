VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries.Interfaces")

Option Explicit

Public Sub ResetIndex()

End Sub

Public Sub AddQuery(ByVal Query As IQuery)

End Sub

Public Function NextQuery() As IQuery

End Function

Public Property Get EndOfQueries() As Boolean

End Property

Public Property Get CurrentIndex() As Long

End Property

Public Property Get RemoveQuery(ByVal Index As Long)

End Property
