'
' 2025/3/4, Aggregate1 Refactoring
'
' Type definition for WellDataForAggregate1
'
Private Type WellDataOne
    Q As Double '양수량
    Q1 As Double '1단계 양수량
    Qg As Double '가채수량
    Qh As Double '한계양수량
    
    Ratio As Double
    
    Sd1 As Double ' 1단계 수위강하령
    Sd2 As Double ' 4단계 수위강하량
    
    C As Double
    B As Double
End Type

' Get well parameters from YangSoo sheet
Private Function GetWellData(wellIndex As Integer) As WellDataOne
    Dim params As WellDataOne
    Dim ws As Worksheet
    Dim row As Long: row = 4 + wellIndex
    
    
    Set ws = Worksheets("YangSoo")

    With params
        .Q = ws.Cells(row, "k").value
        .Qg = ws.Cells(row, "ab").value
        
        .Q1 = ws.Cells(row, "ac").value
        .Qh = ws.Cells(row, "aa").value
        
        .Ratio = ws.Cells(row, "ah").value
        
        .Sd1 = ws.Cells(row, "ad").value
        .Sd2 = ws.Cells(row, "ae").value
        
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
    Dim wd As WellDataOne
    

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

Private Sub WriteWellSummary(WellData As WellDataOne, ByVal wellIndex As Integer, ByVal isSingleWellMode As Boolean)
    ' Writes well summary data to columns G:K and Q:S for a specific well
    ' Parameters:
    '   wellData: Structure containing well measurement data
    '   wellIndex: Index of the well being processed
    '   isSingleWellMode: Flag indicating single well (True) or all wells (False) operation
    
    Dim rowNumber As Integer
    Dim wellLabel As String
    
    ' Calculate target row and well identifier
    rowNumber = wellIndex + 2
    wellLabel = "W-" & wellIndex
    
    ' Clear existing data if in single well mode
    If isSingleWellMode Then
        ClearRange "G" & rowNumber & ":K" & rowNumber
        ClearRange "Q" & rowNumber & ":S" & rowNumber
    End If
    
    ' Write summary data using With blocks for efficiency
    With Range("G" & rowNumber)
        .value = wellLabel
        .Offset(0, 1).value = WellData.Qh
        .Offset(0, 2).value = WellData.Qg
        .Offset(0, 3).value = WellData.Q
        .Offset(0, 4).value = WellData.Ratio
    End With
    
    With Range("Q" & rowNumber)
        .value = wellLabel
        .Offset(0, 1).value = WellData.C
        .Offset(0, 2).value = WellData.B
    End With
    
    ' Apply alternating background formatting
    ApplyBackgroundFormatting rowNumber, "G", "K", wellIndex
    ApplyBackgroundFormatting rowNumber, "Q", "S", wellIndex
End Sub

Private Sub WriteWaterIntake(wd As WellDataOne, ByVal wellIndex As Integer, ByVal isSingleWellMode As Boolean)
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
    Cells(baseRow, "H").value = wd.Sd2
    Cells(baseRow + 1, "H").value = wd.Sd1
    Cells(baseRow, "I").value = wd.Qg

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

