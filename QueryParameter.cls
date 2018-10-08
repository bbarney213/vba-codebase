VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueryParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Queries")

Option Explicit

Private Type TQueryParameter
    PropertyName As String
    MatchValue As Variant
    MatchCondition As MatchCondition
    MatchType As MatchType
    Mode As QueryMode
End Type

Private this As TQueryParameter

Implements IQueryParameter

Private Sub IQueryParameter_Create(ByVal PropertyName As String, ByVal ParameterValue As Variant, ByVal ParameterMatchType As MatchType, ByVal ParameterCondition As MatchCondition)
    With this
        .PropertyName = PropertyName
        
        If IsObject(ParameterValue) Then
            Set .MatchValue = ParameterValue
        Else
            .MatchValue = ParameterValue
        End If
        
        .MatchType = ParameterMatchType
        .MatchCondition = ParameterCondition
    End With
End Sub

Private Property Get IQueryParameter_IsUndefined() As Boolean
    IQueryParameter_IsUndefined = ((this.PropertyName = vbNullString) Or (this.MatchType = mtUndefined) Or (this.MatchCondition = mcUndefined))
End Property

Private Property Get IQueryParameter_PropertyName() As String
    IQueryParameter_PropertyName = this.PropertyName
End Property

Private Property Get IQueryParameter_SatisfiesComparison(ByVal ValueToCompare As Variant) As Boolean
    Dim ModRecordValue As Variant
    ModRecordValue = ModifiedValue(ValueToCompare, this.MatchType)
    
    Dim ModParamValue As Variant
    
    If Not IsObject(this.MatchValue) Then
        ModParamValue = ModifiedValue(this.MatchValue, this.MatchType)
    Else
        Set ModParamValue = this.MatchValue
    End If
    
    IQueryParameter_SatisfiesComparison = Compare(ModRecordValue, ModParamValue, this.MatchCondition)
End Property

Private Property Get IQueryParameter_Value() As Variant
    IQueryParameter_Value = this.MatchValue
End Property

Private Property Get IQueryParameter_MatchType() As MatchType
    IQueryParameter_Value = this.MatchType
End Property

Private Property Get IQueryParameter_Condition() As MatchCondition
    IQueryParameter_Value = this.MatchCondition
End Property

Private Function ModifiedValue(ByVal Value As Variant, ByVal MatchType As MatchType) As Variant
    Select Case MatchType
    Case mtBoolean
        ModifiedValue = CBool(Value)
    Case mtID
        ModifiedValue = Utilities.CID(Value)
    Case mtNumeric
        ModifiedValue = CDbl(Value)
    Case mtText
        ModifiedValue = CStr(Value)
    End Select
End Function

Private Function Compare(ByVal RecordValue As Variant, ByVal ParameterValue As Variant, ByVal MatchCondition As MatchCondition) As Boolean
    Select Case MatchCondition
    Case mcLessThan
        Compare = (RecordValue < ParameterValue)
    Case mcLessThanOrEqualTo
        Compare = (RecordValue <= ParameterValue)
    Case mcEqualTo
        Compare = (RecordValue = ParameterValue)
    Case mcGreaterThanOrEqualTo
        Compare = (RecordValue >= ParameterValue)
    Case mcGreaterThan
        Compare = (RecordValue > ParameterValue)
    Case mcNotEqualTo
        Compare = (RecordValue <> ParameterValue)
    Case mcStartsWith
        Compare = (Left$(RecordValue, Len(ParameterValue)) = ParameterValue)
    Case mcDoesNotStartWith
        Compare = Not (Left$(RecordValue, Len(ParameterValue)) = ParameterValue)
    Case mcEndsWith
        Compare = (Right$(RecordValue, Len(ParameterValue)) = ParameterValue)
    Case mcDoesNotEndWith
        Compare = Not (Right$(RecordValue, Len(ParameterValue)) = ParameterValue)
    Case mcContainsString
        Compare = InStr(1, RecordValue, ParameterValue) > 0
    Case mcDoesNotContainString
        Compare = InStr(1, RecordValue, ParameterValue) = 0
    Case mcIsLike
        Compare = (RecordValue Like ParameterValue)
    Case mcIsNotLike
        Compare = Not (RecordValue Like ParameterValue)
    Case mcInList
        Compare = (ParameterValue.Exists(RecordValue))
    Case mcNotInList
        Compare = Not (ParameterValue.Exists(RecordValue))
    End Select
End Function

   
