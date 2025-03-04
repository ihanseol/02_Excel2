Option Explicit

' Constants for better maintainability
Private Const SUMMARY_START_ROW As Long = 80
Private Const TS_ANALYSIS_START_ROW As Long = 48
Private Const WELL_DATA_START_ROW As Long = 2

' Type definition for well parameters
Private Type WellParameters
    Q As Double
    Natural As Double
    Stable As Double
    Recover As Double
    Radius As Double
    DeltaS As Double
    DeltaH As Double
    DaeSoo As Double
    T1 As Double
    T2 As Double
    TA As Double
    K As Double
    Time As Double
    S1 As Double
    S2 As Double
    Schultz As Double
    Webber As Double
    Jcob As Double
    Skin As Double
    Er As Double
End Type

' Main import procedure
Sub GROK_ImportWellSpec(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim wsAggregate As Worksheet
    Dim wsYangSoo As Worksheet
    Dim wellCount As Integer
    Dim i As Integer
    Dim params As WellParameters

    Set wsAggregate = Worksheets("Aggregate2")
    Set wsYangSoo = Worksheets("YangSoo")
    wsAggregate.Activate

    wellCount = GetNumberOfWell()

    If Not isSingleWellImport Then
        ClearAllDataRanges
    End If

    For i = 1 To wellCount
        If ShouldProcessWell(i, singleWell, isSingleWellImport) Then
            params = GetWellParameters(wsYangSoo, i)

            Application.ScreenUpdating = False
            GROK_WriteAllWellData i, params, isSingleWellImport
            GROK_WriteSummaryTS i, params
            Application.ScreenUpdating = True
        End If
    Next i

    wsAggregate.Range("A1").Select
    Application.CutCopyMode = False
End Sub

' Helper function to determine if well should be processed
Private Function ShouldProcessWell(currentWell As Integer, singleWell As Integer, _
                                 isSingle As Boolean) As Boolean
    ShouldProcessWell = Not isSingle Or (isSingle And currentWell = singleWell)
End Function

' Get well parameters from YangSoo sheet
Private Function GetWellParameters(ws As Worksheet, wellIndex As Integer) As WellParameters
    Dim params As WellParameters
    Dim row As Long: row = 4 + wellIndex

    With params
        .Q = ws.Cells(row, "K").value
        .Natural = ws.Cells(row, "B").value
        .Stable = ws.Cells(row, "C").value
        .Recover = ws.Cells(row, "D").value
        .Radius = ws.Cells(row, "H").value
        .DeltaS = ws.Cells(row, "L").value
        .DeltaH = ws.Cells(row, "F").value
        .DaeSoo = ws.Cells(row, "N").value
        .T1 = ws.Cells(row, "O").value
        .T2 = ws.Cells(row, "P").value
        .TA = ws.Cells(row, "Q").value
        .Time = ws.Cells(row, "U").value
        .S1 = ws.Cells(row, "R").value
        .S2 = ws.Cells(row, "S").value
        .K = ws.Cells(row, "T").value
        .Schultz = ws.Cells(row, "V").value
        .Webber = ws.Cells(row, "W").value
        .Jcob = ws.Cells(row, "X").value
        .Skin = ws.Cells(row, "Y").value
        .Er = ws.Cells(row, "Z").value
    End With

    GetWellParameters = params
End Function

' Write all well data sections
Private Sub GROK_WriteAllWellData(wellIndex As Integer, params As WellParameters, _
                            isSingleImport As Boolean)
    Call GROK_WriteWellData(wellIndex, params, isSingleImport)
    Call GROK_WriteRadiusOfInfluence(wellIndex, params, isSingleImport)
    Call GROK_WriteTSAnalysis(wellIndex, params, isSingleImport)
    Call GROK_WriteRadiusOfInfluence(wellIndex, params, isSingleImport)
    Call GROK_WriteSkinFactor(wellIndex, params, isSingleImport)
    Call GROK_WriteRoiResult(wellIndex, params, isSingleImport)
End Sub

' Write summary T and S values
Private Sub GROK_WriteSummaryTS(wellIndex As Integer, params As WellParameters)
    Dim row As Long: row = SUMMARY_START_ROW + wellIndex - 1

    With Range("H" & row)
        .value = "W-" & wellIndex
        .Offset(0, 1).value = params.T1
        .Offset(0, 2).value = params.S1
    End With
End Sub

' Write well test data (Section 3-3, 3-4, 3-5)
Private Sub GROK_WriteWellData(wellIndex As Integer, params As WellParameters, _
                         isSingleImport As Boolean)
    Dim row As Long: row = WELL_DATA_START_ROW + wellIndex
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "C" & row & ":U" & row
    End If

    With Range("C" & row)
        ' Section 3-3
        .value = "W-" & wellIndex
        .Offset(0, 1).value = 2880
        .Offset(0, 2).value = params.Q
        .Offset(0, 9).value = params.Q
        .Offset(0, 3).value = params.Natural
        .Offset(0, 4).value = params.Stable
        .Offset(0, 5).value = params.Stable - params.Natural
        .Offset(0, 6).value = params.Radius
        .Offset(0, 7).value = params.DeltaS

        ' Section 3-4
        .Offset(0, 10).value = params.Radius
        .Offset(0, 11).value = params.Radius
        .Offset(0, 12).value = params.DaeSoo
        .Offset(0, 13).value = params.T1
        .Offset(0, 14).value = params.S1

        ' Section 3-5
        .Offset(0, 16).value = params.Stable
        .Offset(0, 17).value = params.Recover
        .Offset(0, 18).value = params.Stable - params.Recover
    End With

    ApplyBackgroundFill Range(Cells(row, "C"), Cells(row, "J")), isEven
    ApplyBackgroundFill Range(Cells(row, "L"), Cells(row, "Q")), isEven
    ApplyBackgroundFill Range(Cells(row, "S"), Cells(row, "U")), isEven
End Sub




' WriteData34_skinfactor
Private Sub GROK_WriteSkinFactor(wellIndex As Integer, params As WellParameters, _
                           isSingleImport As Boolean)
    Dim Values As Variant: Values = GetRowColumn("agg2_34_skinfactor")
    Dim baseRow As Long: baseRow = Values(2)
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "P" & baseRow & ":R" & baseRow
    End If

    With Range("P" & (baseRow + wellIndex - 1))
        .value = "W-" & wellIndex
        
        With .Offset(0, 1)
            .value = params.Skin: .NumberFormat = "0.0000"
        End With
        
        With .Offset(0, 2)
            .value = params.Er: .NumberFormat = "0.0000"
        End With
    End With

    ApplyBackgroundFill Range(Cells(baseRow + wellIndex - 1, "P"), Cells(baseRow + wellIndex - 1, "R")), isEven
End Sub


' WriteData38_ROI result
Private Sub GROK_WriteRoiResult(wellIndex As Integer, params As WellParameters, _
                           isSingleImport As Boolean)
    Dim Values As Variant: Values = GetRowColumn("agg2_38_roi_result")
    Dim baseRow As Long: baseRow = Values(2)
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "H" & baseRow & ":N" & baseRow
    End If

    With Range("H" & (baseRow + wellIndex - 1))
        .value = "W-" & wellIndex
        
        With .Offset(0, 1)
            .value = params.Schultz: .NumberFormat = "0.0"
        End With
        
        With .Offset(0, 2)
            .value = params.Webber: .NumberFormat = "0.0"
        End With
        
        
        With .Offset(0, 3)
            .value = params.Jcob: .NumberFormat = "0.0"
        End With
        
        With .Offset(0, 4)
            .value = (params.Schultz + params.Webber + params.Webber) / 3: .NumberFormat = "0.0"
        End With
        
        With .Offset(0, 5)
            .value = Application.WorksheetFunction.max(params.Schultz, params.Webber, params.Webber): .NumberFormat = "0.0"
        End With
        
        With .Offset(0, 6)
            .value = Application.WorksheetFunction.max(params.Schultz, params.Webber, params.Webber): .NumberFormat = "0.0"
        End With
        
    End With

    ApplyBackgroundFill Range(Cells(baseRow + wellIndex - 1, "H"), Cells(baseRow + wellIndex - 1, "N")), isEven
End Sub




' Write radius of influence (Section 3-7)
Private Sub GROK_WriteRadiusOfInfluence(wellIndex As Integer, params As WellParameters, _
                                  isSingleImport As Boolean)
    Dim Values As Variant: Values = GetRowColumn("agg2_37_roi")
    Dim startRow As Long: startRow = Values(2)
    Dim col As Long: col = 3 + wellIndex
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData ColumnNumberToLetter(col) & startRow & ":" & _
                     ColumnNumberToLetter(col) & (startRow + 6)
    End If

    With Cells(startRow, col)
        .Offset(0, 0).value = "W-" & wellIndex
        With .Offset(1, 0)
            .value = params.TA: .NumberFormat = "0.0000"
        End With
        With .Offset(2, 0)
            .value = params.K: .NumberFormat = "0.0000"
        End With
        With .Offset(3, 0)
            .value = params.S2: .NumberFormat = "0.0000000"
        End With
        With .Offset(4, 0)
            .value = params.Time: .NumberFormat = "0.0000"
        End With
        With .Offset(5, 0)
            .value = params.DeltaH: .NumberFormat = "0.00"
        End With
        .Offset(6, 0).value = params.DaeSoo
    End With

    ApplyBackgroundFill Range(Cells(startRow + 1, col), Cells(startRow + 6, col)), isEven
End Sub

' Write TS analysis (Section 3-6)
Private Sub GROK_WriteTSAnalysis(wellIndex As Integer, params As WellParameters, _
                           isSingleImport As Boolean)
    Dim Values As Variant: Values = GetRowColumn("agg2_36_surisangsoo")
    Dim baseRow As Long: baseRow = Values(2) + (wellIndex - 1) * 3
    Dim isEven As Boolean: isEven = (wellIndex Mod 2 = 0)

    If isSingleImport Then
        EraseCellData "C" & baseRow & ":F" & (baseRow + 2)
    End If

    With Range("C" & baseRow)
        .value = "W-" & wellIndex
        .Offset(0, 1).value = "장기양수시험"
        .Offset(1, 1).value = "수위회복시험"
        .Offset(2, 1).value = "선택치"

        With .Offset(0, 2)
            .value = params.T1: .NumberFormat = "0.0000"
        End With
        With .Offset(1, 2)
            .value = params.T2: .NumberFormat = "0.0000"
        End With
        With .Offset(2, 2)
            .value = params.TA: .NumberFormat = "0.0000": .Font.Bold = True
        End With
        With .Offset(0, 3)
            .value = params.S2: .NumberFormat = "0.0000000"
        End With
        With .Offset(2, 3)
            .value = params.S2: .NumberFormat = "0.0000000": .Font.Bold = True
        End With
    End With

    ApplyBackgroundFill Range(Cells(baseRow, "C"), Cells(baseRow, "F")), isEven
End Sub


' Helper sub to clear all data ranges
Private Sub ClearAllDataRanges()
    EraseCellData "C3:J33"
    EraseCellData "L3:Q33"
    EraseCellData "S3:U33"
    EraseCellData "D37:AH43"
    EraseCellData "E48:F137"
    EraseCellData "H48:N77"
    EraseCellData "P48:S77"
    EraseCellData "H80:J109"
End Sub

' Helper sub to apply background fill
Private Sub ApplyBackgroundFill(rng As Range, isEven As Boolean)
    BackGroundFill rng, isEven
End Sub

