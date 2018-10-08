VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries.Interfaces")

Option Explicit

Public Property Get Queries() As QueryCollection

End Property

Public Sub LoadQueriesFromData(ByRef Data As Variant, ByVal PropertyNameHeader As String, _
                                ByVal MatchTypeHeader As String, ByVal MatchConditionHeader As String, _
                                ByVal MatchValueHeader As String, ByVal MatchModeHeader As String, _
                                Optional ByVal Driver As IHeaderDriver = Nothing)

End Sub

