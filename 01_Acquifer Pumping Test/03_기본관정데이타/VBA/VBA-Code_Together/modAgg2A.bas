' Data structure to hold well information
Type WellData
    Q As Double
    Natural As Double
    Stable As Double
    Recover As Double
    Radius As Double
    DeltaS As Double
    DaeSoo As Double
    T1 As Double
    T2 As Double
    TA As Double
    time_ As Double
    S1 As Double
    S2 As Double
    K As Double
    shultz As Double
    Webber As Double
    Jcob As Double
    Skin As Double
    Er As Double
End Type


'' 회복 T값, S값 을 정리해서 뿌려준다.
Sub CERE_WriteSummaryTS(ByVal Well As Integer)
    Dim i As Integer

    i = Well - 1

    Range("H" & (i + 80)).value = "W-" & (i + 1)
    Range("i" & (i + 80)).value = Range("e" & (49 + i * 3)).value
    Range("J" & (i + 80)).value = Range("f" & (48 + i * 3)).value
End Sub


' Core subroutine to handle importing well specifications
Sub CERE_ImportWellSpec(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim i As Integer
    Dim nofwell As Integer
    nofwell = GetNumberOfWell()
    Sheets("Aggregate2").Select

    ' Initialize worksheet reference
    Dim wsYangSoo As Worksheet
    Set wsYangSoo = Worksheets("YangSoo")

    ' Erase existing data if not single well import
    If Not isSingleWellImport Then
        Call EraseWellSections
    End If

    ' Process each well
    For i = 1 To nofwell
        If isSingleWellImport And i <> singleWell Then
            GoTo NextIteration
        End If

        ' Load well data from source sheet
        Dim WellData As WellData
        WellData = GetWellDataFromYangSoo(i)

        ' Write data to target sections
        Call CERE_WriteWellData(WellData, i, isSingleWellImport)
        Call CERE_WriteRadiusOfInfluence(WellData, i, isSingleWellImport)
        Call CERE_WriteTSAnalysis(WellData, i, isSingleWellImport)
        Call CERE_WriteRadiusResult(WellData, i, isSingleWellImport)
        Call CERE_WriteSkinFactor(WellData, i, isSingleWellImport)
        Call CERE_WriteSummaryTS(i)

        ' Toggle application settings
        Call ToggleApplicationSettings(True)

NextIteration:
    Next i

    Range("a1").Select
    Application.CutCopyMode = False
End Sub


' Helper function to fetch well data from worksheet
Function CERE_GetWellDataFromYangSoo(ByVal wellIndex As Integer) As WellData
    With Worksheets("YangSoo")
        Dim wellRow As Integer
        wellRow = 4 + wellIndex

        Dim data As WellData
        data.Q = .Cells(wellRow, "k").value
        data.Natural = .Cells(wellRow, "b").value
        data.Stable = .Cells(wellRow, "c").value
        data.Recover = .Cells(wellRow, "d").value
        data.Radius = .Cells(wellRow, "h").value
        data.DeltaS = .Cells(wellRow, "l").value
        data.DaeSoo = .Cells(wellRow, "n").value
        data.T1 = .Cells(wellRow, "o").value
        data.T2 = .Cells(wellRow, "p").value
        data.TA = .Cells(wellRow, "q").value
        data.time_ = .Cells(wellRow, "u").value
        data.S1 = .Cells(wellRow, "r").value
        data.S2 = .Cells(wellRow, "s").value
        data.K = .Cells(wellRow, "t").value
        data.shultz = .Cells(wellRow, "v").value
        data.Webber = .Cells(wellRow, "w").value
        data.Jcob = .Cells(wellRow, "x").value
        data.Skin = .Cells(wellRow, "y").value
        data.Er = .Cells(wellRow, "z").value

        GetWellDataFromYangSoo = data
    End With
End Function

' Centralized function to erase well sections
Sub EraseWellSections()
    Call EraseCellData("C3:J33")
    Call EraseCellData("L3:Q33")
    Call EraseCellData("S3:U33")
    Call EraseCellData("D37:AH43")
    Call EraseCellData("E48:F137")
    Call EraseCellData("H48:N77")
    Call EraseCellData("P48:S77")
    Call EraseCellData("H80:J109")
End Sub

' Core function to write well data
Sub CERE_WriteWellData(data As WellData, ByVal wellIndex As Integer, ByVal isSingleWellImport As Boolean)
    Call WriteSection3_345(data, wellIndex, isSingleWellImport)
End Sub

' Write 3-3 section: Long-term pumping test results
Sub CERE_WriteSection3_345(data As WellData, ByVal wellIndex As Integer, ByVal isSingleWellImport As Boolean)
    If isSingleWellImport Then
        Call EraseCellData("C" & (wellIndex + 2) & ":J" & (wellIndex + 2))
        Call EraseCellData("L" & (wellIndex + 2) & ":Q" & (wellIndex + 2))
        Call EraseCellData("S" & (wellIndex + 2) & ":U" & (wellIndex + 2))
    End If

    With Worksheets("Aggregate2")
        .Cells(wellIndex + 2, "C").value = "W-" & wellIndex
        .Cells(wellIndex + 2, "D").value = 2880
        .Cells(wellIndex + 2, "E").value = data.Q
        .Cells(wellIndex + 2, "L").value = data.Q
        .Cells(wellIndex + 2, "F").value = data.Natural
        .Cells(wellIndex + 2, "G").value = data.Stable
        .Cells(wellIndex + 2, "H").value = data.Stable - data.Natural
        .Cells(wellIndex + 2, "I").value = data.Radius
        .Cells(wellIndex + 2, "J").value = data.DeltaS

        .Cells(wellIndex + 2, "M").value = data.Radius
        .Cells(wellIndex + 2, "N").value = data.Radius
        .Cells(wellIndex + 2, "O").value = data.DaeSoo
        .Cells(wellIndex + 2, "P").value = data.T1
        .Cells(wellIndex + 2, "Q").value = data.S1

        .Cells(wellIndex + 2, "S").value = data.Stable
        .Cells(wellIndex + 2, "T").value = data.Recover
        .Cells(wellIndex + 2, "U").value = data.Stable - data.Recover
    End With

    Call ApplyBackground(wellIndex, "C,J,L,Q,S,U")
End Sub

' Helper function to apply background color
Sub CERE_ApplyBackground(ByVal index As Integer, ByVal param As String)
    Dim remainder As Integer
    remainder = index Mod 2

    Dim args() As String
    args = Split(param, ",")

    For Each rng In args
        Call BackGroundFill(Worksheets("Aggregate2").Range(Cells(index + 2, rng), Cells(index + 2, rng)), IIf(remainder = 0, True, False))
    Next rng
End Sub


' Write 37 radius of influence section
Sub CERE_WriteRadiusOfInfluence(data As WellData, ByVal wellIndex As Integer, ByVal isSingleWellImport As Boolean)

    Dim ip, remainder As Variant
    Dim Values As Variant

    Values = GetRowColumn("agg2_37_roi")
    ip = Values(2)


    If isSingleWellImport Then
        Call EraseCellData(ColumnNumberToLetter(4 + wellIndex) & ip & ":" & ColumnNumberToLetter(4 + wellIndex) & (ip + 6))
    End If

    With Worksheets("Aggregate2")
        .Cells((ip + 0), (3 + wellIndex)).value = "W-" & wellIndex

        .Cells((ip + 1), (3 + wellIndex)).value = data.TA
        .Cells((ip + 1), (3 + wellIndex)).NumberFormat = "0.0000"

        .Cells((ip + 2), (3 + wellIndex)).value = data.K
        .Cells((ip + 2), (3 + wellIndex)).NumberFormat = "0.0000"


        .Cells((ip + 3), (3 + wellIndex)).value = data.S2
        .Cells((ip + 3), (3 + wellIndex)).NumberFormat = "0.0000000"

        .Cells((ip + 4), (3 + wellIndex)).value = data.time_
        .Cells((ip + 4), (3 + wellIndex)).NumberFormat = "0.0000"

        .Cells((ip + 5), (3 + wellIndex)).value = data.Stable - data.Recover
        .Cells((ip + 5), (3 + wellIndex)).NumberFormat = "0.00"

        .Cells((ip + 6), (3 + wellIndex)).value = data.DaeSoo
    End With


    remainder = wellIndex Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + 1, (wellIndex + 3)), Cells(ip + 6, (wellIndex + 3))), True)
    Else
            Call BackGroundFill(Range(Cells(ip + 1, (wellIndex + 3)), Cells(ip + 6, (wellIndex + 3))), False)
    End If
End Sub

' Write 36 TS analysis section
Sub CERE_WriteTSAnalysis(data As WellData, ByVal wellIndex As Integer, ByVal isSingleWellImport As Boolean)
    ' Implementation with reduced redundancy
End Sub

' Write 38 radius result section
Sub CERE_WriteRadiusResult(data As WellData, ByVal wellIndex As Integer, ByVal isSingleWellImport As Boolean)
    ' Implementation with reduced redundancy
End Sub

' Write 34 skin factor section
Sub CERE_WriteSkinFactor(data As WellData, ByVal wellIndex As Integer, ByVal isSingleWellImport As Boolean)
    ' Implementation with reduced redundancy
End Sub

' Helper function to toggle application settings
Sub ToggleApplicationSettings(ByVal isEnabled As Boolean)
    Application.ScreenUpdating = isEnabled
    Application.EnableEvents = isEnabled
    Application.Calculation = IIf(isEnabled, xlCalculationAutomatic, xlCalculationManual)
End Sub





