Sub SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range
    Dim dataRanges As Object
    Set dataRanges = CreateObject("Scripting.Dictionary")

    ' Set references to worksheets
    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    ' Define data ranges for each dataArrayName
    With dataRanges
        .Add "Q", wsInput.Range("m51")
        .Add "hp", wsInput.Range("i48")
        .Add "natural", wsInput.Range("m48")
        .Add "stable", wsInput.Range("m49")
        .Add "radius", wsInput.Range("m44")
        .Add "Rw", wsSkinFactor.Range("e4")
        .Add "well_depth", wsInput.Range("m45")
        .Add "casing", wsInput.Range("i52")
        .Add "C", wsInput.Range("A31")
        .Add "B", wsInput.Range("B31")
        .Add "recover", wsSkinFactor.Range("c10")
        .Add "Sw", wsSkinFactor.Range("c11")
        .Add "delta_h", wsSkinFactor.Range("b16")
        .Add "delta_s", wsSkinFactor.Range("b4")
        .Add "daeSoo", wsSkinFactor.Range("c16")
        .Add "T0", wsSkinFactor.Range("d4")
        .Add "S0", wsSkinFactor.Range("f4")
        .Add "ER_MODE", wsSkinFactor.Range("h10")
        .Add "T1", wsSkinFactor.Range("d5")
        .Add "T2", wsSkinFactor.Range("h13")
        .Add "TA", wsSkinFactor.Range("d16")
        .Add "S1", wsSkinFactor.Range("e10")
        .Add "S2", wsSkinFactor.Range("i16")
        .Add "K", wsSkinFactor.Range("e16")
        .Add "time_", wsSkinFactor.Range("h16")
        .Add "shultze", wsSkinFactor.Range("c13")
        .Add "webber", wsSkinFactor.Range("c18")
        .Add "jacob", wsSkinFactor.Range("c23")
        .Add "skin", wsSkinFactor.Range("g6")
        .Add "er", wsSkinFactor.Range("c8")
        .Add "ER1", wsSkinFactor.Range("k8")
        .Add "ER2", wsSkinFactor.Range("k9")
        .Add "ER3", wsSkinFactor.Range("k10")
        .Add "qh", wsSafeYield.Range("b13")
        .Add "qg", wsSafeYield.Range("b7")
        .Add "sd1", wsSafeYield.Range("b3")
        .Add "sd2", wsSafeYield.Range("b4")
        .Add "q1", wsSafeYield.Range("b2")
        .Add "ratio", wsSafeYield.Range("b11")
    End With

    ' Set dataCell based on dataArrayName
    If dataRanges.Exists(dataArrayName) Then
        Set dataCell = dataRanges(dataArrayName)
        SetCellValueForWell wellIndex, dataCell, dataArrayName
    Else
        MsgBox "Data array name not found: " & dataArrayName
    End If
End Sub


'*****************************************************************************

Sub SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range

    
    Dim dataRanges() As Variant
    Dim addresses() As Variant
    Dim i As Integer

    ' Set references to worksheets
    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    ' Define data ranges for each dataArrayName
    dataRanges = Array(wsInput.Range("m51"), wsInput.Range("i48"), _
                        wsInput.Range("m48"), wsInput.Range("m49"), _
                        wsInput.Range("m44"), wsSkinFactor.Range("e4"), _
                        wsInput.Range("m45"), wsInput.Range("i52"), _
                        wsInput.Range("A31"), wsInput.Range("B31"), _
                        wsSkinFactor.Range("c10"), wsSkinFactor.Range("c11"), _
                        wsSkinFactor.Range("b16"), wsSkinFactor.Range("b4"), _
                        wsSkinFactor.Range("c16"), wsSkinFactor.Range("d4"), _
                        wsSkinFactor.Range("f4"), wsSkinFactor.Range("h10"), _
                        wsSkinFactor.Range("d5"), wsSkinFactor.Range("h13"), _
                        wsSkinFactor.Range("d16"), wsSkinFactor.Range("e10"), _
                        wsSkinFactor.Range("i16"), wsSkinFactor.Range("e16"), _
                        wsSkinFactor.Range("h16"), wsSkinFactor.Range("c13"), _
                        wsSkinFactor.Range("c18"), wsSkinFactor.Range("c23"), _
                        wsSkinFactor.Range("g6"), wsSkinFactor.Range("c8"), _
                        wsSkinFactor.Range("k8"), wsSkinFactor.Range("k9"), _
                        wsSkinFactor.Range("k10"), wsSafeYield.Range("b13"), _
                        wsSafeYield.Range("b7"), wsSafeYield.Range("b3"), _
                        wsSafeYield.Range("b4"), wsSafeYield.Range("b2"), _
                        wsSafeYield.Range("b11"))

    ' Array of data addresses
    addresses = Array("Q", "hp", "natural", "stable", "radius", "Rw", _
                        "well_depth", "casing", "C", "B", "recover", "Sw", _
                        "delta_h", "delta_s", "daeSoo", "T0", "S0", "ER_MODE", _
                        "T1", "T2", "TA", "S1", "S2", "K", "time_", "shultze", _
                        "webber", "jacob", "skin", "er", "ER1", "ER2", "ER3", _
                        "qh", "qg", "sd1", "sd2", "q1", "ratio")

    ' Find index of dataArrayName in addresses array
    For i = LBound(addresses) To UBound(addresses)
        If addresses(i) = dataArrayName Then
            Set dataCell = dataRanges(i)
            Exit For
        End If
    Next i

    ' Check if dataArrayName is found
    If Not dataCell Is Nothing Then
        SetCellValueForWell wellIndex, dataCell, dataArrayName
    Else
        MsgBox "Data array name not found: " & dataArrayName
    End If
End Sub



'*****************************************************************************



Sub SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range
    Dim value As Variant

    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    Select Case dataArrayName
        Case "Q"
            Set dataCell = wsInput.Range("m51")
        Case "hp"
            Set dataCell = wsInput.Range("i48")
        
        
        Case "natural"
            Set dataCell = wsInput.Range("m48")
        Case "stable"
            Set dataCell = wsInput.Range("m49")
        Case "radius"
            Set dataCell = wsInput.Range("m44")
        Case "Rw"
            Set dataCell = wsSkinFactor.Range("e4")
        
        Case "well_depth"
            Set dataCell = wsInput.Range("m45")
        Case "casing"
            Set dataCell = wsInput.Range("i52")
        
        Case "C"
            Set dataCell = wsInput.Range("A31")
         Case "B"
            Set dataCell = wsInput.Range("B31")
        
        
        Case "recover"
            Set dataCell = wsSkinFactor.Range("c10")
        Case "Sw"
            Set dataCell = wsSkinFactor.Range("c11")
        
        Case "delta_h"
            Set dataCell = wsSkinFactor.Range("b16")
        Case "delta_s"
            Set dataCell = wsSkinFactor.Range("b4")
    
        Case "daeSoo"
            Set dataCell = wsSkinFactor.Range("c16")
            
  '--------------------------------------------------------------
  
       Case "T0"
            Set dataCell = wsSkinFactor.Range("d4")
        Case "S0"
            Set dataCell = wsSkinFactor.Range("f4")
       Case "ER_MODE"
            Set dataCell = wsSkinFactor.Range("h10")
                  
        Case "T1"
            Set dataCell = wsSkinFactor.Range("d5")
        Case "T2"
            Set dataCell = wsSkinFactor.Range("h13")
        Case "TA"
            Set dataCell = wsSkinFactor.Range("d16")
            
       Case "S1"
            Set dataCell = wsSkinFactor.Range("e10")
        Case "S2"
            Set dataCell = wsSkinFactor.Range("i16")
        
        Case "K"
            Set dataCell = wsSkinFactor.Range("e16")
        Case "time_"
            Set dataCell = wsSkinFactor.Range("h16")
            
        Case "shultze"
            Set dataCell = wsSkinFactor.Range("c13")
        Case "webber"
            Set dataCell = wsSkinFactor.Range("c18")
        Case "jacob"
            Set dataCell = wsSkinFactor.Range("c23")
                    
                        
       Case "skin"
            Set dataCell = wsSkinFactor.Range("g6")
        Case "er"
            Set dataCell = wsSkinFactor.Range("c8")
            
        Case "ER1"
            Set dataCell = wsSkinFactor.Range("k8")
        Case "ER2"
            Set dataCell = wsSkinFactor.Range("k9")
        Case "ER3"
            Set dataCell = wsSkinFactor.Range("k10")


        Case "qh"
            Set dataCell = wsSafeYield.Range("b13")
        Case "qg"
            Set dataCell = wsSafeYield.Range("b7")
            
        Case "sd1"
            Set dataCell = wsSafeYield.Range("b3")
        Case "sd2"
            Set dataCell = wsSafeYield.Range("b4")
        Case "q1"
            Set dataCell = wsSafeYield.Range("b2")
        Case "ratio"
            Set dataCell = wsSafeYield.Range("b11")
    End Select

    SetCellValueForWell wellIndex, dataCell, dataArrayName
End Sub


'****************************************************************************************************************************************************************************************************************


Another method to refactor the SetDataArrayValues subroutine is to utilize a custom data structure to store the mappings between data array names and their corresponding ranges. One option is to use a class module to define this data structure. Here's how you can do it:

Create a Class Module:
First, insert a new Class Module into your VBA project. You can name it DataRangeMapping.

Define Properties:
Within the class module, define properties for the data array name and its corresponding range. You can name them DataArrayName and DataRange, respectively.

Code for Class Module:
Your class module code should look something like this:

vb
Copy code
' Class module: DataRangeMapping

Public DataArrayName As String
Public DataRange As Range


Refactor the Subroutine:
Now, you can refactor the SetDataArrayValues subroutine to use instances of this class to store the mappings. Here's how you can refactor the subroutine:

vb
Copy code


Sub SetDataArrayValues(ByVal wb As Workbook, ByVal wellIndex As Integer, ByVal dataArrayName As String)
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataCell As Range
    Dim dataMappings As New Collection
    Dim mapping As DataRangeMapping
    Dim foundMapping As Boolean

    ' Set references to worksheets
    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    ' Define data range mappings
    Set mapping = New DataRangeMapping
    With mapping
        .DataArrayName = "Q"
        Set .DataRange = wsInput.Range("m51")
    End With
    dataMappings.Add mapping

    Set mapping = New DataRangeMapping
    With mapping
        .DataArrayName = "hp"
        Set .DataRange = wsInput.Range("i48")
    End With
    dataMappings.Add mapping

    ' Add more mappings for other data array names and ranges...

    ' Find the mapping for the given dataArrayName
    For Each mapping In dataMappings
        If mapping.DataArrayName = dataArrayName Then
            Set dataCell = mapping.DataRange
            foundMapping = True
            Exit For
        End If
    Next mapping

    ' Set the cell value if mapping is found
    If foundMapping Then
        SetCellValueForWell wellIndex, dataCell, dataArrayName
    Else
        MsgBox "Data array name not found: " & dataArrayName
    End If
End Sub

Update SetCellValueForWell:
You may need to adjust the SetCellValueForWell subroutine to accept a range parameter instead of a cell parameter, depending on how it's defined.

This method provides a more scalable and maintainable solution by encapsulating the data range mappings within a class module. It also makes it easier to add or modify mappings without directly modifying the subroutine code.



'****************************************************************************************************************************************************************************************************************

