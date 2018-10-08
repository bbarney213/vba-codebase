VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries")

Option Explicit

Private Type TQuery
    QueryMode As QueryMode
    Parameters As Collection
End Type

Private this As TQuery

Implements IQuery

Private Sub Class_Initialize()
    Set this.Parameters = New Collection
End Sub

Private Sub IQuery_AddParameter(ByVal ParameterName As String, ByVal ParameterValue As Variant, ByVal ParameterMatchType As MatchType, ByVal ParameterCondition As MatchCondition)
    Dim Parameter As IQueryParameter
    Set Parameter = New QueryParameter
    
    Parameter.Create ParameterName, ParameterValue, ParameterMatchType, ParameterCondition
    
    If Not Parameter.IsUndefined Then
        this.Parameters.Add Parameter
    End If
End Sub

Public Property Get IQuery_Parameters() As Collection
    Set IQuery_Parameters = this.Parameters
End Property

Private Property Get IQuery_Mode() As QueryMode
    IQuery_Mode = this.QueryMode
End Property

Private Property Let IQuery_Mode(Value As QueryMode)
    this.QueryMode = Value
End Property

Private Property Get IQuery_SatisfiesParameters(ByVal Record As IQueryableRecord) As Boolean
    If this.QueryMode = qmAndMode Then
        IQuery_SatisfiesParameters = SatisfiesAndMode(Record)
    Else
        IQuery_SatisfiesParameters = SatisfiesOrMode(Record)
    End If
End Property

Private Function SatisfiesAndMode(ByVal Record As IQueryableRecord) As Boolean
    Dim Parameter As IQueryParameter
    For Each Parameter In this.Parameters
        If Not Parameter.SatisfiesComparison(Record.GetNamedProperty(Parameter.PropertyName)) Then
            Exit Function
        End If
    Next
    SatisfiesAndMode = True
End Function
Private Function SatisfiesOrMode(ByVal Record As IQueryableRecord) As Boolean
    Dim Parameter As IQueryParameter
    For Each Parameter In this.Parameters
        If Parameter.SatisfiesComparison(Record.GetNamedProperty(Parameter.PropertyName)) Then
            SatisfiesOrMode = True
            Exit Function
        End If
    Next
End Function

