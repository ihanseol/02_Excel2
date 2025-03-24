
Sub WriteStepTestData(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
' isSingleWellImport = True ---> SingleWell Import
' isSingleWellImport = False ---> AllWell Import
'
' SingleWell --> ImportWell Number
' 999 & False --> 모든관정을 임포트
'

    Dim nofwell, i As Integer
    Dim a1, a2, a3, Q, h, delta_h, qsw, swq As String
    Dim fName As String
    
    nofwell = GetNumberOfWell()
    
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim rngString As String
    
    
    Call TurnOffStuff
    
    If ActiveSheet.name <> "AggStep" Then Sheets("AggStep").Select
    
    
    If isSingleWellImport Then
        rngString = "C" & (singleWell + 5 - 1) & ":K" & (singleWell + 5 - 1)
        Call EraseCellData(rngString)
    Else
        rngString = "C5:K36"
        
        fName = "A1_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Call EraseCellData(rngString)
    End If
        
    
    For i = 1 To nofwell
    
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If
    
SINGLE_ITERATION:

        fName = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Set wb = Workbooks(fName)
        Set wsInput = wb.Worksheets("Input")
        
        Q = wsInput.Range("q64").value
        h = wsInput.Range("r64").value
        delta_h = wsInput.Range("s64").value
        qsw = wsInput.Range("t64").value
        swq = wsInput.Range("u64").value

        a1 = wsInput.Range("v64").value
        a2 = wsInput.Range("w64").value
        a3 = wsInput.Range("x64").value
        
        Call Write31_StepTestData_Single(a1, a2, a3, Q, h, delta_h, qsw, swq, i)

NEXT_ITERATION:

    Next i
    
    Call TurnOnStuff
    'Call Write31_StepTestData(a1, a2, a3, Q, h, delta_h, qsw, swq, nofwell)
End Sub


Sub Write31_StepTestData_Single(a1 As Variant, a2 As Variant, a3 As Variant, Q As Variant, h As Variant, delta_h As Variant, qsw As Variant, swq As Variant, i As Integer)
' i : well_index
    
    ' Call EraseCellData("C5:K36")
    
    Cells(4 + i, "c").value = "W-" & CStr(i)
    
    Cells(4 + i, "d").value = a1
    Cells(4 + i, "e").value = a2
    Cells(4 + i, "f").value = a3

    Cells(4 + i, "g").value = Q
    Cells(4 + i, "h").value = h
    Cells(4 + i, "i").value = delta_h
    Cells(4 + i, "j").value = qsw
    Cells(4 + i, "k").value = swq

End Sub

