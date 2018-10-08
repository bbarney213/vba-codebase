VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Version 1.1.1

'@Folder("CodeBase.Classes.Recordset")

Option Explicit

Private Type TRecordCollection
    Records As Collection
End Type

Private this As TRecordCollection

Implements IQueryableRecords
Implements IRecordCollection

Private Sub Class_Initialize()
    Set this.Records = New Collection
End Sub

Private Sub IQueryableRecords_AddRecord(ByVal Record As IQueryableRecord)
    this.Records.Add Record
End Sub

Private Property Get IQueryableRecords_FilteredRecords(ByVal Queries As IQueryCollection) As Collection
    Dim FilteredRecords As Collection
    Set FilteredRecords = New Collection
    
    Dim Record As IQueryableRecord
    For Each Record In this.Records
        FilteredRecords.Add Record
    Next
    
    If FilteredRecords.Count > 0 Then
        Do Until Queries.EndOfQueries
            Dim Query As IQuery
            Set Query = Queries.NextQuery
        
            Dim i As Long
            For i = FilteredRecords.Count To 1 Step -1
                If Not Query.SatisfiesParameters(FilteredRecords(i)) Then
                    FilteredRecords.Remove i
                End If
            Next
        Loop
        
        Queries.ResetIndex
    End If
    
    Set IQueryableRecords_FilteredRecords = FilteredRecords
End Property

Private Property Set IQueryableRecords_Records(value As Collection)
    Set this.Records = value
End Property

Private Property Get IQueryableRecords_Records() As Collection
    Set IQueryableRecords_Records = this.Records
End Property

Public Property Get IRecordCollection_Records() As Collection
    Set IRecordCollection_Records = this.Records
End Property

Private Property Set IRecordCollection_Records(value As Collection)
    Set this.Records = value
End Property

Private Sub IRecordCollection_AddRecord(ByVal Record As IAnonymousRecord)
    this.Records.Add Record
End Sub



