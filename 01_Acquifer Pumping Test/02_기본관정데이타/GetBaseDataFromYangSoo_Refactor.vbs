Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim nofwell As Integer
    Dim i As Integer
    Dim rngString As String

    ' Arrays to store data
    Dim dataArrays As Variant
    dataArrays = Array("natural", "stable", "recover", "delta_h", "Sw", "radius", _
                       "Rw", "well_depth", "casing", "Q", "delta_s", "hp", _
                       "daeSoo", "T1", "T2", "TA", "S1", "S2", "K", "time_", _
                       "shultze", "webber", "jacob", "skin", "er", "ER1", _
                       "ER2", "ER3", "qh", "qg", "sd1", "sd2", "q1", "C", _
                       "B", "ratio", "T0", "S0", "ER_MODE")

    ' Check if all well data should be imported
    nofwell = GetNumberOfWell()
    If Not isSingleWellImport And singleWell = 999 Then
        rngString = "A5:AN" & (nofwell + 5 - 1)
        Call EraseCellData(rngString)
    End If

    ' Loop through each well
    For i = 1 To nofwell
        ' Import data for all wells or only for the specified single well
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            ImportDataForWell i, dataArrays
        End If
    Next i
End Sub

Sub ImportDataForWell(ByVal wellIndex As Integer, ByVal dataArrays As Variant)
    Dim fName As String
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    Dim dataIdx As Integer
    Dim cellOffset As Integer
    Dim dataCell As Range

    ' Open the workbook
    fName = "A" & CStr(wellIndex) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data! " & fName
        Exit Sub
    End If
    Set wb = Workbooks(fName)

    ' Loop through data arrays and import values
    For dataIdx = LBound(dataArrays) To UBound(dataArrays)
        SetDataArrayValues wb, wellIndex, dataArrays(dataIdx)
    Next dataIdx

    ' Set additional values not in the arrays
    Set wsInput = wb.Worksheets("Input")
    Set wsSkinFactor = wb.Worksheets("SkinFactor")
    Set wsSafeYield = wb.Worksheets("SafeYield")

    With wsSkinFactor
        Set dataCell = .Range("d4")
        SetCellValueForWell wellIndex, dataCell, "T0"

        Set dataCell = .Range("f4")
        SetCellValueForWell wellIndex, dataCell, "S0"

        Set dataCell = .Range("h10")
        SetCellValueForWell wellIndex, dataCell, "ER_MODE"

        Set dataCell = .Range("d5")
        SetCellValueForWell wellIndex, dataCell, "T1"

        Set dataCell = .Range("h13")
        SetCellValueForWell wellIndex, dataCell, "T2"

        Set dataCell = .Range("e10")
        SetCellValueForWell wellIndex, dataCell, "S1"

        Set dataCell = .Range("i16")
        SetCellValueForWell wellIndex, dataCell, "S2"

        ' Add more cells here as needed
    End With

    ' Close workbook
    wb.Close SaveChanges:=False
End Sub

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
        Case "natural"
            Set dataCell = wsInput.Range("m48")
        Case "stable"
            Set dataCell = wsInput.Range("m49")
        Case "recover"
            Set dataCell = wsSkinFactor.Range("c10")
        ' Add more cases for other arrays
        ' ...
    End Select

    SetCellValueForWell wellIndex, dataCell, dataArrayName
End Sub

Sub SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim wellData As Double

    wellData = dataCell.Value
    Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).Value = wellData
    If dataArrayName = "recover" Or dataArrayName = "Sw" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "S2" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000000"
    ElseIf dataArrayName = "T1" Or dataArrayName = "T2" Or dataArrayName = "TA" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    ElseIf dataArrayName = "qh" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0."
    ElseIf dataArrayName = "qg" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "q1" Or dataArrayName = "sd1" Or dataArrayName = "sd2" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "skin" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).Value = Format(wellData, "0.0000")
    ElseIf dataArrayName = "er" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    ElseIf dataArrayName = "ratio" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0%"
    ElseIf dataArrayName = "T0" Or dataArrayName = "S0" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    End If
End Sub

Function GetColumnIndex(ByVal columnName As String) As Integer
    Dim colIndex As Integer

    Select Case columnName
        Case "natural"
            colIndex = 2
        Case "stable"
            colIndex = 3
        Case "recover"
            colIndex = 4
        ' Add more cases for other column names
        ' ...
    End Select

    GetColumnIndex = colIndex
End Function
