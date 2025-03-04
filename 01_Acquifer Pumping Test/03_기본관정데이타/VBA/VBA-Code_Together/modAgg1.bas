'
' 2025/3/4, Aggregate1 Refactoring
'

' Type definition for WellDataForAggregate1
'
Private Type WellDataForAggOne
    Q1 As Double
    QQ1 As Double
    Q2 As Double
    Q3 As Double
    Ratio As Double
    
    S1 As Double
    S2 As Double
    
    C As Double
    B As Double
End Type

' Get well parameters from YangSoo sheet
Private Function GetWellData(wellIndex As Integer) As WellDataForAggOne
    Dim params As WellDataForAggOne
    Dim ws As Worksheet
    Dim row As Long: row = 4 + wellIndex
    
    
    Set ws = Worksheets("YangSoo")

    With params
        .Q1 = ws.Cells(row, "aa").value
        .QQ1 = ws.Cells(row, "ac").value
        
        .Q2 = ws.Cells(row, "ab").value
        .Q3 = ws.Cells(row, "k").value
        
        .Ratio = ws.Cells(row, "ah").value
        
        .S1 = ws.Cells(row, "ad").value
        .S2 = ws.Cells(row, "ae").value
        
        .C = ws.Cells(row, "af").value
        .B = ws.Cells(row, "ag").value
    End With

    GetWellData = params
End Function


Sub ImportAggregateData(ByVal targetWell As Integer, ByVal isSingleWellMode As Boolean)
    ' Handles both single well and all wells import operations
    ' isSingleWellMode = True: Imports data for specified well only
    ' isSingleWellMode = False: Imports data for all wells

    Dim wellCount As Integer
    Dim wellIndex As Integer
    Dim wd As WellDataForAggOne
    

    ' Initialize core variables
    wellCount = GetNumberOfWell()
    
    Sheets("Aggregate1").Activate

    ' Clear data ranges if importing all wells
    If Not isSingleWellMode Then
        ClearRange "G3:K35"
        ClearRange "Q3:S35"
        ClearRange "F43:I102"
    End If

    ' Process each well
    For wellIndex = 1 To wellCount
        If ShouldProcessWell(wellIndex, targetWell, isSingleWellMode) Then
            ' Fetch well data from YangSoo worksheet
           
            wd = GetWellData(wellIndex)
            
            ' Process data with optimizations disabled
            TurnOffStuff
            
            Call WriteWellSummary(wd, wellIndex, isSingleWellMode)
            Call WriteWaterIntake(wd, wellIndex, isSingleWellMode)
            
            TurnOnStuff
        End If
    Next wellIndex

    ' Clean up
    Application.CutCopyMode = False
    Range("L1").Select
End Sub

Private Sub WriteWellSummary(wd As WellDataForAggOne, ByVal wellIndex As Integer, ByVal isSingleWellMode As Boolean)
    ' Writes well summary data to G:K and Q:S ranges

    Dim rowNumber As Integer
    Dim i As Integer
    Dim wellLabel As String

    rowNumber = wellIndex + 2
    wellLabel = "W-" & wellIndex

    ' Clear specific row if in single well mode
    If isSingleWellMode Then
        ClearRange "G" & rowNumber & ":K" & rowNumber
        ClearRange "Q" & rowNumber & ":S" & rowNumber
    End If
    
    i = wellIndex

    ' Write data to summary columns (G:K)
    Range("G" & (i + 2)).value = "W-" & i
    Range("H" & (i + 2)).value = wd.Q1
    Range("I" & (i + 2)).value = wd.Q2
    Range("J" & (i + 2)).value = wd.Q3
    Range("K" & (i + 2)).value = wd.Ratio

    Range("Q" & (i + 2)).value = "W-" & i
    Range("R" & (i + 2)).value = wd.C
    Range("S" & (i + 2)).value = wd.B
    
    
    ' Apply background formatting
    ApplyBackgroundFormatting rowNumber, "G", "K", wellIndex
    ApplyBackgroundFormatting rowNumber, "Q", "S", wellIndex
End Sub

Private Sub WriteWaterIntake(wd As WellDataForAggOne, ByVal wellIndex As Integer, ByVal isSingleWellMode As Boolean)
    ' Calculates and writes tentative water intake data starting at row 43

    Dim startRow As Integer
    Dim baseRow As Integer
    Dim values As Variant

    ' Get starting row from configuration
    values = GetRowColumn("Agg1_Tentative_Water_Intake")
    startRow = values(2)
    baseRow = startRow + (wellIndex - 1) * 2

    ' Clear specific rows if in single well mode
    If isSingleWellMode Then
        ClearRange "F" & baseRow & ":I" & (baseRow + 1)
    End If

    ' Write water intake data
    Cells(baseRow, "F").value = "W-" & CStr(wellIndex)
    Cells(baseRow, "G").value = wd.Q1
    Cells(baseRow, "H").value = wd.S2
    Cells(baseRow + 1, "H").value = wd.S1
    Cells(baseRow, "I").value = wd.Q2

    ' Apply background formatting
    ApplyBackgroundFormatting baseRow, "F", "I", wellIndex, 2
End Sub

Private Function ShouldProcessWell(ByVal currentIndex As Integer, ByVal targetWell As Integer, _
                                 ByVal isSingleWellMode As Boolean) As Boolean
    ' Determines if a well should be processed based on import mode
    ShouldProcessWell = Not isSingleWellMode Or (isSingleWellMode And currentIndex = targetWell)
End Function

Private Sub ApplyBackgroundFormatting(ByVal startRow As Integer, ByVal startCol As String, _
                                    ByVal endCol As String, ByVal wellIndex As Integer, _
                                    Optional ByVal rowSpan As Integer = 1)
    ' Applies alternating background colors to specified range
    Dim targetRange As Range
    Set targetRange = Range(Cells(startRow, startCol), Cells(startRow + rowSpan - 1, endCol))
    BackGroundFill targetRange, (wellIndex Mod 2 = 0)
End Sub

Private Sub ClearRange(ByVal rangeAddress As String)
    ' Clears content in specified range
    Range(rangeAddress).ClearContents
End Sub

