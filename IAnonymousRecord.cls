VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAnonymousRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.3

'@Folder("CodeBase.Classes.Anonymous.Interfaces")

Option Explicit

Public Sub AddNamedProperty(ByVal PropertyName As String, ByVal PropertyValue As Variant, Optional ByVal UpdateIfExisting As Boolean)

End Sub
Public Function GetNamedProperty(ByVal PropertyName As String) As Variant
    
End Function
Public Sub RemoveNamedProperty(ByVal PropertyName As String)

End Sub
Public Sub UpdateNamedProperty(ByVal PropertyName As String, ByVal PropertyValue As Variant)

End Sub

Public Property Get NamedProperties() As Scripting.Dictionary

End Property

