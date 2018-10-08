VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries")

Option Explicit

Private Const QUERY_MODE_CODE_AND As String = "AND"
Private Const QUERY_MODE_CODE_OR As String = "OR"

Private Type TQueryFactory
    Queries As IQueryCollection
End Type

Private this As TQueryFactory

Implements IQueryFactory

Private Sub IQueryFactory_LoadQueriesFromData(ByRef Data As Variant, ByVal PropertyNameHeader As String, _
                                                ByVal MatchTypeHeader As String, ByVal MatchConditionHeader As String, _
                                                ByVal MatchValueHeader As String, ByVal MatchModeHeader As String, _
                                                Optional ByVal Driver As IHeaderDriver = Nothing)
    Set this.Queries = New QueryCollection
    
    Dim HeaderIndices As Scripting.Dictionary
    Set HeaderIndices = Utilities.GetHeaderIndices(Data)
    
    If Not HeadersExist(Array(PropertyNameHeader, MatchTypeHeader, MatchConditionHeader, MatchValueHeader, MatchModeHeader), HeaderIndices) Then
        Err.Raise ERR_NUMBER_HEADERS_MISSING, Description:=ERR_MESSAGE_HEADERS_MISSING
    End If
    
    ' How do I best, dynamically, solve this?
    Dim AndModeQuery As IQuery
    Set AndModeQuery = New Query
    AndModeQuery.Mode = qmAndMode
    
    Dim OrModeQuery As IQuery
    Set OrModeQuery = New Query
    OrModeQuery.Mode = qmOrMode
        
    Dim i As Long
    For i = LBound(Data, 1) + 1 To UBound(Data, 1)
        Dim Property As String
        Property = Data(i, HeaderIndices(PropertyNameHeader))
        
        If Not Driver Is Nothing Then
            Property = Driver.NewHeader(Property)
        End If
        
        Dim MatchType As MatchType
        MatchType = ConvertToMatchType(Data(i, HeaderIndices(MatchTypeHeader)))
        
        Dim MatchCondition As MatchCondition
        MatchCondition = ConvertToMatchCondition(Data(i, HeaderIndices(MatchConditionHeader)))
        
        Dim MatchValue As Variant
        MatchValue = Data(i, HeaderIndices(MatchValueHeader))
        
        Select Case Data(i, HeaderIndices(MatchModeHeader))
        Case QUERY_MODE_CODE_AND
            AndModeQuery.AddParameter Property, MatchValue, MatchType, MatchCondition
        Case QUERY_MODE_CODE_OR
            OrModeQuery.AddParameter Property, MatchValue, MatchType, MatchCondition
        Case vbNullString
        Case Else
            Err.Raise ERR_NUMBER_FUNCTIONALITY_NOT_IMPLEMENTED, Description:=ERR_MESSAGE_FUNCTIONALITY_NOT_IMPLEMENTED
        End Select
    Next
    
    this.Queries.AddQuery AndModeQuery
    this.Queries.AddQuery OrModeQuery
End Sub

Private Property Get IQueryFactory_Queries() As QueryCollection
    Set IQueryFactory_Queries = this.Queries
End Property

Private Function ConvertToMatchType(ByVal InputValue As String) As MatchType
    Select Case InputValue
    
    Case QUERIES_MATCH_TYPE_BOOLEAN
        ConvertToMatchType = mtBoolean
        
    Case QUERIES_MATCH_TYPE_NUMERIC
        ConvertToMatchType = mtNumeric
        
    Case QUERIES_MATCH_TYPE_TEXT
        ConvertToMatchType = mtText
        
    Case QUERIES_MATCH_TYPE_ID
        ConvertToMatchType = mtID
        
    Case Else
        ConvertToMatchType = mtUndefined
    End Select
End Function
Private Function ConvertToMatchCondition(ByVal InputValue) As MatchCondition
    Select Case InputValue
    Case QUERIES_MATCH_CONDITION_EQUAL_TO
        ConvertToMatchCondition = mcEqualTo
        
    Case QUERIES_MATCH_CONDITION_NOT_EQUAL_TO
        ConvertToMatchCondition = mcNotEqualTo
        
    Case QUERIES_MATCH_CONDITION_LESS_THAN
        ConvertToMatchCondition = mcLessThan
        
    Case QUERIES_MATCH_CONDITION_LESS_THAN_OR_EQUAL_TO
        ConvertToMatchCondition = mcLessThanOrEqualTo
        
    Case QUERIES_MATCH_CONDITION_GREATER_THAN
        ConvertToMatchCondition = mcGreaterThan
        
    Case QUERIES_MATCH_CONDITION_GREATER_THAN_OR_EQUAL_TO
        ConvertToMatchCondition = mcGreaterThanOrEqualTo
        
    Case QUERIES_MATCH_CONDITION_STARTS_WITH
        ConvertToMatchCondition = mcStartsWith
        
    Case QUERIES_MATCH_CONDITION_ENDS_WITH
        ConvertToMatchCondition = mcEndsWith
        
    Case QUERIES_MATCH_CONDITION_CONTAINS
        ConvertToMatchCondition = mcContainsString
        
    Case QUERIES_MATCH_CONDITION_DOES_NOT_CONTAIN
        ConvertToMatchCondition = mcDoesNotContainString
        
    Case QUERIES_MATCH_CONDITION_IS_LIKE
        ConvertToMatchCondition = mcIsLike
        
    Case Else
        ConvertToMatchCondition = mcUndefined
    End Select
End Function




















