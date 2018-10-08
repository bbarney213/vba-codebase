Attribute VB_Name = "DriverTemplate"
' Version 1.1.1

'@Folder("CodeBase.Utilities.Drivers")

Option Explicit

Sub PrintDataHeaderCodeFromSelection()
    Dim Data As Variant
    Data = Selection.value
    
    Debug.Print "Dim Foo as Scripting.Dictionary"
    Debug.Print "Set Foo = New Scripting.Dictionary"
    Debug.Print vbNewLine
    Debug.Print "With"
    
    Dim i As Long
    For i = LBound(Data, 2) To UBound(Data, 2)
        Debug.Print vbTab & ".Add " & Chr(34) & Data(LBound(Data, 1), i) & Chr(34) & ", " & Chr(34) & Data(UBound(Data, 1), i) & Chr(34)
    Next
    Debug.Print "End With"
    
End Sub

Sub PrintRecordHeaderCodeFromSelection()
    Dim Data As Variant
    Data = Selection.value
    
    Debug.Print "Dim Foo as Scripting.Dictionary"
    Debug.Print "Set Foo = New Scripting.Dictionary"
    Debug.Print vbNewLine
    Debug.Print "With"
    
    Dim i As Long
    For i = LBound(Data, 2) To UBound(Data, 2)
        Debug.Print vbTab & ".Add " & Chr(34) & Data(UBound(Data, 1), i) & Chr(34) & ", " & Chr(34) & Data(LBound(Data, 1), i) & Chr(34)
    Next
    Debug.Print "End With"
    
End Sub

