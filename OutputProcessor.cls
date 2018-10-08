VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutputProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CodeBase.Classes.OutputProcessor")

Option Explicit

Private Const MAX_WORKSHEET_ROWS As Long = 1048576
Private Const NAME_SEPARATOR As String = "__"
Private Const DEFAULT_NAME As String = "Output__Processor"

Private Type TOutputProcessor
    Name As String
    
    CurrentOutputWorkbook As Workbook
    CurrentOutputWorksheet As Worksheet
    CurentOutputTable As ListObject
    
    OutputWorksheets As Collection
    
    OutputHeaders As Scripting.Dictionary
End Type

Private this As TOutputProcessor

Implements IOutputProcessor

Private Sub Class_Initialize()
    Set this.OutputWorksheets = New Collection
End Sub

Private Property Let IOutputProcessor_Name(value As String)
    this.Name = value
End Property

Private Property Get IOutputProcessor_OutputHeaders() As Scripting.IDictionary
    Set IOutputProcessor_OutputHeaders = this.OutputHeaders
End Property

Private Property Set IOutputProcessor_OutputHeaders(value As Scripting.IDictionary)
    Set this.OutputHeaders = value
End Property

Private Property Get IOutputProcessor_OutputWorkbook() As Workbook
    Set IOutputProcessor_OutputWorkbook = this.CurrentOutputWorkbook
End Property

Private Property Set IOutputProcessor_OutputWorkbook(value As Workbook)
    Set this.CurrentOutputWorkbook = value
End Property

Private Property Get IOutputProcessor_OutputWorksheet() As Worksheet
    Set IOutputProcessor_OutputWorksheet = this.CurrentOutputWorksheet
End Property

Private Property Set IOutputProcessor_OutputWorksheet(value As Worksheet)
    Set this.CurrentOutputWorksheet = value
End Property

Private Sub IOutputProcessor_OutputDataUsingFactory(ByVal Factory As IAnonymousFactory, Optional ByVal HeadersToUse As Scripting.Dictionary = Nothing)
    If this.OutputHeaders Is Nothing Then LoadHeaders Factory, HeadersToUse
    
    If Factory.Records.Count > 0 Then
        OutputData Factory.Data(True, this.OutputHeaders)
    End If
End Sub

Private Sub LoadHeaders(ByVal Factory As IAnonymousFactory, ByVal HeadersToUse As Scripting.Dictionary)
    If HeadersToUse Is Nothing Then
        Set this.OutputHeaders = Factory.Headers
    Else
        Set this.OutputHeaders = HeadersToUse
    End If
End Sub

Private Sub OutputData(ByRef Data As Variant)
    If this.Name = vbNullString Then this.Name = DEFAULT_NAME
    
    If NotEnoughRowsAvailable(Data) Or this.CurrentOutputWorksheet Is Nothing Then
        LoadNewOutputWorksheet
    End If
    
    With this.CurentOutputTable
        If .DataBodyRange Is Nothing Then
            .Range(.ListRows.Count + 2, 1).Resize((UBound(Data, 1) - LBound(Data, 1) + 1), .ListColumns.Count).value = Data
        Else
            .DataBodyRange(.ListRows.Count + 1, 1).Resize((UBound(Data, 1) - LBound(Data, 1) + 1), .ListColumns.Count).value = Data
        End If
    End With
End Sub

Private Function NotEnoughRowsAvailable(ByRef Data As Variant) As Boolean
    If Not this.CurrentOutputWorksheet Is Nothing Then
        NotEnoughRowsAvailable = (this.CurrentOutputWorksheet.UsedRange.Rows.Count + NumberOfArrayRecords(Data)) > MAX_WORKSHEET_ROWS
    End If
End Function

Private Function NumberOfArrayRecords(ByRef Data As Variant) As Long
    If IsArray(Data) Then
        NumberOfArrayRecords = UBound(Data, 1) - LBound(Data, 1) + 1
    End If
End Function

Private Sub PrepareOutputWorkbook()
    Dim Worksheet As Worksheet
    Dim DisplayAlerts As Boolean
    DisplayAlerts = Application.DisplayAlerts
            
    Dim ScreenUpdating As Boolean
    ScreenUpdating = Application.ScreenUpdating
            
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    For Each Worksheet In this.CurrentOutputWorkbook.Worksheets
        If Worksheet.Name Like this.Name & "__*" Then
            Worksheet.Delete
        End If
    Next
    
    Application.DisplayAlerts = DisplayAlerts
    Application.ScreenUpdating = ScreenUpdating
End Sub

Private Sub LoadNewOutputWorksheet()
    If this.CurrentOutputWorkbook Is Nothing Then Set this.CurrentOutputWorkbook = ThisWorkbook
    
    Dim OutputWorksheet As Worksheet
    With this.CurrentOutputWorkbook
        Set OutputWorksheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        
        Dim OutputName As String
        OutputName = this.Name & NAME_SEPARATOR & this.OutputWorksheets.Count + 1
        
        On Error Resume Next
        .Worksheets(OutputName).Visible = xlSheetVisible
        .Worksheets(OutputName).Delete
        On Error GoTo 0
        
        OutputWorksheet.Name = OutputName
        LoadNewOutputTable OutputWorksheet
        
        this.OutputWorksheets.Add OutputWorksheet
        
        Set this.CurrentOutputWorksheet = OutputWorksheet
    End With
End Sub

Private Sub LoadNewOutputTable(ByVal Worksheet As Worksheet)
    Dim Headers As Variant
    ReDim Headers(0 To 0, 0 To this.OutputHeaders.Count - 1)
    
    Dim Header As Variant
    For Each Header In this.OutputHeaders.Keys
        Dim i As Long
        Headers(0, i) = Header
        i = i + 1
    Next
    
    With Worksheet
        .Cells.Clear
        
        Dim OutputRange As Range
        Set OutputRange = .Range("A1").Resize((UBound(Headers, 1) - LBound(Headers, 1) + 1), (UBound(Headers, 2) - LBound(Headers, 2)) + 1)
        
        OutputRange.value = Headers
        
        Dim OutputTable As ListObject
        Set OutputTable = .ListObjects.Add(xlSrcRange, OutputRange, , xlYes)
        
        With OutputTable
            .Name = Worksheet.Name
            
            With .Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End With
        
        Set this.CurentOutputTable = OutputTable
    End With
End Sub

Private Function ProcessOutputData(ByRef Data As Variant, ByVal HeaderIndices As Scripting.Dictionary) As Variant
    Dim Start As Double
    Start = Timer
    
    Dim InputIndices As Scripting.Dictionary
    Set InputIndices = Utilities.GetHeaderIndices(Data)
    
    Dim OutputData As Variant
    ReDim OutputData(LBound(Data, 1) To UBound(Data, 1) - 1, LBound(Data, 2) To UBound(Data, 2))
    
    Dim i As Long
    For i = LBound(OutputData, 1) To UBound(OutputData, 1)
        Dim Header As Variant
        For Each Header In HeaderIndices
            If InputIndices.Exists(Header) Then
                OutputData(i, HeaderIndices(Header)) = Data(i + 1, InputIndices(Header))
            End If
        Next
        
        DoEvents
    Next

    ProcessOutputData = OutputData
    
    Debug.Print Round(Timer - Start, 2) & " seconds to Process Output Data"
End Function
