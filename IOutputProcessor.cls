VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOutputProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CodeBase.Classes.OutputProcessor.Interfaces")

Option Explicit

Public Property Let Name(Value As String)

End Property

Public Property Get OutputWorkbook() As Workbook

End Property

Public Property Set OutputWorkbook(Value As Workbook)

End Property

Public Property Set OutputWorksheet(Value As Worksheet)

End Property

Public Property Get OutputWorksheet() As Worksheet

End Property

Public Property Get OutputHeaders() As Scripting.Dictionary

End Property

Public Property Set OutputHeaders(Value As Scripting.Dictionary)

End Property

Public Sub OutputDataUsingFactory(ByVal Factory As IAnonymousFactory, Optional ByVal HeadersToUse As Scripting.Dictionary = Nothing)

End Sub
