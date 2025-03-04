Attribute VB_Name = "modWell_Each"
Sub ImportEachWell(ByVal well_no As Integer)
    ' well_no -- well number
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    Dim nl, sl, DeltaS As Double
    Dim casing As Integer
    Dim wsYangSoo As Worksheet

    ' Set the well number
    i = well_no
    
    ' Reference the YangSoo worksheet
    Set wsYangSoo = Worksheets("YangSoo")
    
    ' Turn off additional processes or features
    BaseData_ETC_02.TurnOffStuff
    
    ' Read data from the worksheet
    DeltaS = wsYangSoo.Cells(4 + i, "L").value
    nl = wsYangSoo.Cells(4 + i, "B").value
    sl = wsYangSoo.Cells(4 + i, "C").value
    casing = wsYangSoo.Cells(4 + i, "J").value
    T1 = wsYangSoo.Cells(4 + i, "O").value
    S1 = wsYangSoo.Cells(4 + i, "R").value
    T2 = wsYangSoo.Cells(4 + i, "P").value
    S2 = wsYangSoo.Cells(4 + i, "S").value
    S3 = wsYangSoo.Cells(4 + i, "AQ").value
    Skin = wsYangSoo.Cells(4 + i, "Y").value
    RI1 = wsYangSoo.Cells(4 + i, "V").value
    RI2 = wsYangSoo.Cells(4 + i, "W").value
    RI3 = wsYangSoo.Cells(4 + i, "X").value
    
    ' Calculate the effective radius
    ir = GetEffectiveRadiusFromFX(i)
    
    ' Set the values in the target worksheet
    SetCellValueAndFormat Range("C20"), nl, "0.00"
    SetCellValueAndFormat Range("C21"), sl, "0.00"
    SetCellValueAndFormat Range("C10"), 5, "0"
    SetCellValueAndFormat Range("C11"), casing - 5, "0"
    SetCellValueAndFormat Range("G6"), S3, "0.00"
    SetCellValueAndFormat Range("E5"), T1, "0.0000"
    SetCellValueAndFormat Range("E6"), T2, "0.0000"
    SetCellValueAndFormat Range("G5"), S2, "0.0000000"
    SetCellValueAndFormat Range("G4"), S1, "0.00000"
    SetCellValueAndFormat Range("H5"), Skin, "0.0000"
    SetCellValueAndFormat Range("H6"), ir, "0.0000"
    SetCellValueAndFormat Range("E10"), RI1, "0.0"
    SetCellValueAndFormat Range("F10"), RI2, "0.0"
    SetCellValueAndFormat Range("G10"), RI3, "0.0"
    SetCellValueAndFormat Range("C23"), Round(DeltaS, 2), "0.00"
    
    ' Turn on additional processes or features
    BaseData_ETC_02.TurnOnStuff
End Sub

' Helper function to set cell value and format
Sub SetCellValueAndFormat(cell As Range, value As Variant, format As String)
    cell.value = value
    cell.NumberFormat = format
End Sub


Sub ImportWellSpecFX(ByVal well_no As Integer)
'
' well_no -- well number
'
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    
    ' nl : natural level, sl : stable level
    ' s3 - Recover Test 의 S값
    
    Dim nl, sl, DeltaS As Double
    Dim casing As Integer
    Dim wsYangSoo As Worksheet
    
    i = well_no
    Set wsYangSoo = Worksheets("YangSoo")
    BaseData_ETC_02.TurnOffStuff
    
    ' delta s : 최초1분의 수위강하
    DeltaS = wsYangSoo.Cells(4 + i, "L").value
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    nl = wsYangSoo.Cells(4 + i, "B").value
    sl = wsYangSoo.Cells(4 + i, "C").value
    casing = wsYangSoo.Cells(4 + i, "J").value
    
    
    T1 = wsYangSoo.Cells(4 + i, "O").value
    S1 = wsYangSoo.Cells(4 + i, "R").value
    T2 = wsYangSoo.Cells(4 + i, "P").value
    S2 = wsYangSoo.Cells(4 + i, "S").value
    S3 = wsYangSoo.Cells(4 + i, "AQ").value
    
    ' 스킨계수
    Skin = wsYangSoo.Cells(4 + i, "Y").value
    
    ' yangsoo radius of influence
    RI1 = wsYangSoo.Cells(4 + i, "V").value  ' schultze
    RI2 = wsYangSoo.Cells(4 + i, "W").value  ' webber
    RI3 = wsYangSoo.Cells(4 + i, "X").value  ' jcob
    
    ' 유효우물반경 , 설정값에 따른
    ' ir = GetEffectiveRadius(WBNAME)
    ir = GetEffectiveRadiusFromFX(i)
    
      ' Set the values in the target worksheet
    SetCellValueAndFormat Range("C20"), nl, "0.00"
    SetCellValueAndFormat Range("C21"), sl, "0.00"
    SetCellValueAndFormat Range("C10"), 5, "0"
    SetCellValueAndFormat Range("C11"), casing - 5, "0"
    SetCellValueAndFormat Range("G6"), S3, "0.00"
    SetCellValueAndFormat Range("E5"), T1, "0.0000"
    SetCellValueAndFormat Range("E6"), T2, "0.0000"
    SetCellValueAndFormat Range("G5"), S2, "0.0000000"
    SetCellValueAndFormat Range("G4"), S1, "0.00000"
    SetCellValueAndFormat Range("H5"), Skin, "0.0000"
    SetCellValueAndFormat Range("H6"), ir, "0.0000"
    SetCellValueAndFormat Range("E10"), RI1, "0.0"
    SetCellValueAndFormat Range("F10"), RI2, "0.0"
    SetCellValueAndFormat Range("G10"), RI3, "0.0"
    SetCellValueAndFormat Range("C23"), Round(DeltaS, 2), "0.00"

    BaseData_ETC_02.TurnOnStuff

End Sub





Private Sub ImportEachWell_OLD()
    Dim WkbkName As Object
    Dim WBNAME, cell1 As String
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, Skin As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, DeltaS As Double
    Dim casing As Integer
    
    BaseData_ETC_02.TurnOffStuff
    
    i = 2
    ' Range("i1") = Workbooks.count
    ' WBName = Range("i2").value
    
    cell1 = Range("b2").value
    WBNAME = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBNAME) Then
        MsgBox "Please open the yangsoo data ! " & WBNAME
        Exit Sub
    End If

    ' delta s : 최초1분의 수위강하
    DeltaS = Workbooks(WBNAME).Worksheets("SkinFactor").Range("b4").value
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    nl = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i4").value
    sl = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i6").value
    casing = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i10").value
    
    ' WkbkName.Close
    T1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("D5").value
    S1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("E10").value
    T2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("H13").value
    S2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i16").value
    S3 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i13").value
    
    Skin = Workbooks(WBNAME).Worksheets("SkinFactor").Range("G6").value
    
    ' yangsoo radius of influence
    ' 슐츠, 영향반경
    RI1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C13").value
    ' 웨버, 영향반경
    RI2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C18").value
    ' 제이콥, 영향반경
    RI3 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C23").value
    
    ' 유효우물반경 , 설정값에 따른
    ir = GetEffectiveRadius(WBNAME)
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    Range("c20") = nl
    Range("c20").NumberFormat = "0.00"
    
    Range("c21") = sl
    Range("c21").NumberFormat = "0.00"
    
    Range("c10") = 5
    Range("c11") = casing - 5
    
    'in recover test, s' value
    Range("G6") = S3
        
    Range("E5") = T1
    Range("E5").NumberFormat = "0.0000"
     
    Range("E6") = T2
    Range("E6").NumberFormat = "0.0000"
    
    Range("g5") = S2
    Range("g5").NumberFormat = "0.0000000"
    
    '2024/6/10 move to s1 this G4 cell
    Range("G4") = S1
    
    
    Range("h5") = Skin 'skin coefficient
    Range("h6") = ir    'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(DeltaS, 2) 'deltas
    
    BaseData_ETC_02.TurnOnStuff
        
End Sub
