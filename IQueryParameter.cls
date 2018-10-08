VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IQueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries.Interfaces")

Option Explicit

Public Sub Create(ByVal PropertyName As String, ByVal ParameterValue As Variant, ByVal ParameterMatchType As MatchType, ByVal ParameterCondition As MatchCondition)

End Sub

Public Property Get IsUndefined() As Boolean

End Property

Public Property Get PropertyName() As String

End Property

Public Property Get Value() As Variant

End Property

Public Property Get MatchType() As MatchType

End Property

Public Property Get Condition() As MatchCondition

End Property

Public Property Get SatisfiesComparison(ByVal ValueToCompare As Variant) As Boolean

End Property
