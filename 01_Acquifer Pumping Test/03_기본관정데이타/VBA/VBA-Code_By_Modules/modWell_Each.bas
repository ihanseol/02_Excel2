Attribute VB_Name = "modWell_Each"

Sub ImporEachWell(ByVal well_no As Integer)
'
' well_no -- well number
'
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, skin As Double
    
    ' nl : natural level, sl : stable level
    ' s3 - Recover Test 의 S값
    
    Dim nl, sl, deltas As Double
    Dim casing As Integer
    Dim wsYangSoo As Worksheet
    
    i = well_no
    Set wsYangSoo = Worksheets("YangSoo")
    BaseData_ETC_02.TurnOffStuff
    
    ' delta s : 최초1분의 수위강하
    deltas = wsYangSoo.Cells(4 + i, "L").value
    
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
    skin = wsYangSoo.Cells(4 + i, "Y").value
    
    ' yangsoo radius of influence
    RI1 = wsYangSoo.Cells(4 + i, "V").value  ' schultze
    RI2 = wsYangSoo.Cells(4 + i, "W").value  ' webber
    RI3 = wsYangSoo.Cells(4 + i, "X").value  ' jcob
    
    ' 유효우물반경 , 설정값에 따른
    ' ir = GetEffectiveRadius(WBNAME)
    ir = GetEffectiveRadiusFromFX(i)
    
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
    
    
    Range("h5") = skin 'skin coefficient
    Range("h6") = ir    'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(deltas, 2) 'deltas
    BaseData_ETC_02.TurnOnStuff

End Sub

