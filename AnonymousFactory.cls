VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnonymousFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.2

'@Folder("CodeBase.Classes.Anonymous")

Option Explicit

Private Type TAnonymousFactory
    RecordCollection As IRecordCollection
    FilteredRecordSet As IRecordCollection
    Driver As IHeaderDriver
End Type

Private this As TAnonymousFactory

Implements IAnonymousFactory
Implements IQueryableFactory

Private Sub Class_Initialize()
    Set this.RecordCollection = New RecordCollection
End Sub

Public Sub IAnonymousFactory_LoadRecordsFromData(ByRef InputData As Variant, Optional ByRef RecordDriver As IHeaderDriver = Nothing)
    LoadRecordsFromData InputData, RecordDriver
End Sub

Private Sub IAnonymousFactory_AddGlobalProperty(ByRef PropertyName As String, ByRef PropertyValue As Variant)
    Dim Record As IAnonymousRecord
    For Each Record In this.RecordCollection.Records
        Record.AddNamedProperty PropertyName, PropertyValue
    Next
End Sub

Private Property Get IAnonymousFactory_Records() As Collection
    Set IAnonymousFactory_Records = this.RecordCollection.Records
End Property

Private Property Set IAnonymousFactory_Records(ByRef value As Collection)
    Set this.RecordCollection.Records = value
End Property

Private Property Get IAnonymousFactory_Data(Optional ByRef ExcludeHeaders As Boolean = False, Optional ByRef OutputHeadersToUse As Scripting.Dictionary) As Variant
    Dim OutputHeaders As Scripting.Dictionary
    If OutputHeadersToUse Is Nothing Then
        Set OutputHeaders = GetOutputHeaders
    Else
        Set OutputHeaders = OutputHeadersToUse
    End If
    
    IAnonymousFactory_Data = GetData(ExcludeHeaders, OutputHeaders)
End Property

Private Property Get IAnonymousFactory_Headers() As Scripting.IDictionary
    Set IAnonymousFactory_Headers = GetOutputHeaders
End Property

Private Sub IQueryableFactory_LoadRecordsFromData(ByRef InputData As Variant, Optional ByRef RecordDriver As IHeaderDriver = Nothing)
                                                    
    LoadRecordsFromData InputData, RecordDriver
End Sub

Private Property Get IQueryableFactory_Data(Optional ByRef ExcludeHeaders As Boolean = False, Optional ByRef OutputHeadersToUse As Scripting.Dictionary = Nothing) As Variant
    Dim OutputHeaders As Scripting.Dictionary
    If OutputHeadersToUse Is Nothing Then
        Set OutputHeaders = GetOutputHeaders
    Else
        Set OutputHeaders = OutputHeadersToUse
    End If
    
    IQueryableFactory_Data = GetData(ExcludeHeaders, OutputHeaders)
End Property

Private Property Get IQueryableFactory_FilteredRecords(ByRef Queries As IQueryCollection) As IRecordCollection
    Dim Records As IQueryableRecords
    Set Records = this.RecordCollection
    
    Dim FilteredRecords As IRecordCollection
    Set FilteredRecords = New RecordCollection
    
    Set FilteredRecords.Records = Records.FilteredRecords(Queries)
    
    Set this.FilteredRecordSet = FilteredRecords
    
    Set IQueryableFactory_FilteredRecords = FilteredRecords
End Property

Private Property Get IQueryableFactory_Records() As Collection
    Set IQueryableFactory_Records = this.RecordCollection.Records
End Property

Private Property Set IQueryableFactory_Records(value As Collection)
    Set this.RecordCollection.Records = value
End Property

' Can this be abstracted a level further by becoming a responsibility of the container?
'   If following the conventions elsewhere, an AnonymousFactory should be responsible for creation and not containing.
Private Sub LoadRecordsFromData(ByRef InputData As Variant, Optional ByRef RecordDriver As IHeaderDriver)
    Dim HeaderIndices As Scripting.Dictionary
    Set HeaderIndices = Utilities.GetHeaderIndices(InputData)
    
    If Not RecordDriver Is Nothing Then
        Set HeaderIndices = GetConvertedRecordHeaders(HeaderIndices, RecordDriver)
    End If
    
    Dim i As Long
    For i = LBound(InputData, 1) + 1 To UBound(InputData, 1)
       Dim Record As IAnonymousRecord
       Set Record = New AnonymousRecord
        
        With Record
            Dim Header As Variant
            For Each Header In HeaderIndices.Keys
                .AddNamedProperty Header, InputData(i, HeaderIndices(Header))
            Next
        End With
        
        this.RecordCollection.AddRecord Record
        DoEvents
    Next
End Sub

Private Function GetData(ByRef ExcludeHeaders As Boolean, ByRef HeadersToUse As Scripting.Dictionary) As Variant
    If this.RecordCollection.Records.Count > 0 Then
        Dim OutputData As Variant
        ReDim OutputData(0 To this.RecordCollection.Records.Count - (CLng(ExcludeHeaders) * -1), 0 To HeadersToUse.Count - 1)
        
        If Not ExcludeHeaders Then
            Dim Property As Variant
            For Each Property In HeadersToUse.Keys
                OutputData(0, HeadersToUse(Property)) = Property
            Next
        End If
        
        Dim i As Long
        If Not ExcludeHeaders Then i = 1
        
        Dim Record As IAnonymousRecord
        For Each Record In this.RecordCollection.Records
            For Each Property In HeadersToUse.Keys
                OutputData(i, HeadersToUse(Property)) = Record.GetNamedProperty(Property)
            Next
            i = i + 1
            
            DoEvents
        Next
        
        GetData = OutputData
    End If
End Function

Private Function GetOutputHeaders() As Scripting.Dictionary
    If this.RecordCollection.Records.Count > 0 Then
        Dim SampleRecord As IAnonymousRecord
        Set SampleRecord = this.RecordCollection.Records(1)
        
        Dim Headers As Scripting.Dictionary
        Set Headers = New Scripting.Dictionary
        
        Dim Property As Variant
        For Each Property In SampleRecord.NamedProperties
            Headers.Add Property, Headers.Count
        Next
        
        Set GetOutputHeaders = Headers
    End If
End Function

Private Function GetConvertedRecordHeaders(ByRef InputHeaderIndices As Scripting.Dictionary, ByRef ConversionDriver As IHeaderDriver) As Scripting.Dictionary
    Dim OutputHeaderIndices As Scripting.Dictionary
    Set OutputHeaderIndices = New Scripting.Dictionary
    
    Dim Header As Variant
    For Each Header In InputHeaderIndices.Keys
        Dim ConvertedHeader As String
        ConvertedHeader = ConversionDriver.NewHeader(Header)
        
        If ConvertedHeader <> vbNullString Then
            If Not OutputHeaderIndices.Exists(ConvertedHeader) Then
                OutputHeaderIndices.Add ConvertedHeader, InputHeaderIndices(Header)
            End If
        End If
    Next
    
    Set GetConvertedRecordHeaders = OutputHeaderIndices
End Function


