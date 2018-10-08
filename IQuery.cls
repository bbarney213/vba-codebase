VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries.Interfaces")

Option Explicit

Public Enum QueryMode
    qmOrMode
    qmAndMode
    qmUndefined
End Enum

Public Enum MatchType
    mtBoolean
    mtNumeric
    mtText
    mtID
    mtUndefined
End Enum

Public Enum MatchCondition
    mcEqualTo
    mcNotEqualTo
    mcLessThan
    mcLessThanOrEqualTo
    mcGreaterThan
    mcGreaterThanOrEqualTo
    mcStartsWith
    mcEndsWith
    mcContainsString
    mcDoesNotContainString
    mcIsLike
    mcInList
    mcUndefined
End Enum

Public Property Get SatisfiesParameters(ByVal Record As IQueryableRecord) As Boolean

End Property

Public Property Get Parameters() As Collection

End Property

Public Sub AddParameter(ByVal ParameterName As String, ByVal ParameterValue As Variant, ByVal ParameterMatchType As MatchType, ByVal ParameterCondition As MatchCondition)

End Sub

Public Property Get Mode() As QueryMode

End Property

Public Property Let Mode(Value As QueryMode)

End Property
