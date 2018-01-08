VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnonymousRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.2

'@Folder("CodeBase.Classes.Anonymous")

Private Const DATA_HEADER_COLLEAGUE_ID = "ColleagueID"

Option Explicit

Private Type TAnonymousRecord
    NamedProperties As Scripting.Dictionary
End Type

Implements IAnonymousRecord
Implements IQueryableRecord

Private this As TAnonymousRecord

Private Sub Class_Initialize()
    Set this.NamedProperties = New Scripting.Dictionary
End Sub

Public Sub IAnonymousRecord_AddNamedProperty(ByVal PropertyName As String, ByVal PropertyValue As Variant)
    With this.NamedProperties
        If Not .Exists(PropertyName) Then
            If PropertyName = DATA_HEADER_COLLEAGUE_ID Then
                .Add PropertyName, Utilities.CID(PropertyValue)
            Else
                .Add PropertyName, PropertyValue
            End If
        End If
    End With
End Sub

Public Function IAnonymousRecord_GetNamedProperty(ByVal PropertyName As String) As Variant
    If this.NamedProperties.Exists(PropertyName) Then
        IAnonymousRecord_GetNamedProperty = this.NamedProperties(PropertyName)
    End If
End Function

Public Sub IAnonymousRecord_RemoveNamedProperty(ByVal PropertyName As String)
    If this.NamedProperties.Exists(PropertyName) Then
        this.NamedProperties.Remove PropertyName
    End If
End Sub

Public Sub IAnonymousRecord_UpdateNamedProperty(ByVal PropertyName As String, ByVal PropertyValue As Variant)
    If this.NamedProperties.Exists(PropertyName) Then
        If IsObject(PropertyValue) Then
            Set this.NamedProperties(PropertyName) = PropertyValue
        Else
            this.NamedProperties(PropertyName) = PropertyValue
        End If
    End If
End Sub

Public Property Get IAnonymousRecord_NamedProperties() As Scripting.Dictionary
    Set IAnonymousRecord_NamedProperties = this.NamedProperties
End Property

Public Function IQueryableRecord_GetNamedProperty(ByVal PropertyName As String) As Variant
    If this.NamedProperties.Exists(PropertyName) Then
        IQueryableRecord_GetNamedProperty = this.NamedProperties(PropertyName)
    End If
End Function

