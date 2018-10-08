VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHeaderDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Drivers.Interfaces")

Option Explicit

Public Enum DriverState
    Historical = 0
    Current = 1
End Enum

Public Sub ChangeDriverState(ByVal Setting As DriverState)

End Sub

Public Property Get NewHeader(ByVal InputHeader As String) As String

End Property

Public Property Get OldHeader(ByVal InputHeader As String) As String

End Property
