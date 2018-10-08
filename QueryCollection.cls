VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries")

Option Explicit

Private Type TQueryCollection
    Index As Long
    Queries As Collection
End Type

Private this As TQueryCollection

Implements IQueryCollection

Private Sub Class_Initialize()
    Set this.Queries = New Collection
End Sub

Private Sub IQueryCollection_AddQuery(ByVal Query As IQuery)
    If Not Query.Parameters Is Nothing Then
        If Query.Parameters.Count > 0 Then
            this.Queries.Add Query
        End If
    End If
End Sub

Private Property Get IQueryCollection_CurrentIndex() As Long
    IQueryCollection_CurrentIndex = this.Index
End Property

Private Function IQueryCollection_NextQuery() As IQuery
    If Not IQueryCollection_EndOfQueries Then
        this.Index = this.Index + 1
        Set IQueryCollection_NextQuery = this.Queries(this.Index)
    End If
End Function

Private Property Get IQueryCollection_EndOfQueries() As Boolean
    IQueryCollection_EndOfQueries = (this.Index >= this.Queries.Count)
End Property

Private Property Get IQueryCollection_RemoveQuery(ByVal Index As Long) As Variant
    If Not Index > this.Queries.Count Then
        If this.Index > Index Then this.Index = this.Index - 1
        
        this.Queries.Remove Index
    End If
End Property

Private Sub IQueryCollection_ResetIndex()
    this.Index = 0
End Sub
