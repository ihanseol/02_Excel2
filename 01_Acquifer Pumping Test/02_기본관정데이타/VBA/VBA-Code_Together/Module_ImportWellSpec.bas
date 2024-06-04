
Sub ImportWellSpec(ByVal well_no As Integer)
    Dim WkbkName As Object
    Dim WBName As String
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, deltas As Double
    Dim casing As Integer

    WBName = "A" & GetNumeric2(well_no) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBName) Then
        MsgBox "Please open the yangsoo data ! " & WBName
        Exit Sub
    End If

    ' delta s : 최초1분의 수위강하
    deltas = Workbooks(WBName).Worksheets("SkinFactor").Range("b4").value
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    nl = Workbooks(WBName).Worksheets("SkinFactor").Range("i4").value
    sl = Workbooks(WBName).Worksheets("SkinFactor").Range("i6").value
    casing = Workbooks(WBName).Worksheets("SkinFactor").Range("i10").value
    
    ' WkbkName.Close
    T1 = Workbooks(WBName).Worksheets("SkinFactor").Range("D5").value
    S1 = Workbooks(WBName).Worksheets("SkinFactor").Range("E10").value
    T2 = Workbooks(WBName).Worksheets("SkinFactor").Range("H13").value
    S2 = Workbooks(WBName).Worksheets("SkinFactor").Range("i16").value
    S3 = Workbooks(WBName).Worksheets("SkinFactor").Range("i13").value
    
    ' yangsoo radius of influence
    RI1 = Workbooks(WBName).Worksheets("SkinFactor").Range("C13").value
    RI2 = Workbooks(WBName).Worksheets("SkinFactor").Range("C18").value
    RI3 = Workbooks(WBName).Worksheets("SkinFactor").Range("C23").value
    
    ' 유효우물반경 , 설정값에 따른
    ir = GetEffectiveRadius(WBName)
    
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
    
    Range("h7") = S1
    
    
    Range("h6") = ir 'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(deltas, 2) 'deltas
End Sub

