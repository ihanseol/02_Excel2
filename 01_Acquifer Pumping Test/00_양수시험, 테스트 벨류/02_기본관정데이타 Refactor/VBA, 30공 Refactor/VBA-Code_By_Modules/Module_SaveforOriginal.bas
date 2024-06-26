Attribute VB_Name = "Module_SaveforOriginal"
' Sheet_AggFX, Module Save
'
'Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
'' isSingleWellImport = True ---> SingleWell Import
'' isSingleWellImport = False ---> AllWell Import
''
'' SingleWell --> ImportWell Number
'' 999 & False --> 모든관정을 임포트
''
'    Dim fName As String
'    Dim nofwell, i As Integer
'    Dim rngString As String
'
'    Dim natural() As Double ' 자연수위, natural depth
'    Dim stable() As Double  ' 안정수위, stable depth
'    Dim recover() As Double ' 회복수위, recover depth
'    Dim Sw() As Double ' 수위회복량 - 안정수위 - 회복수위
'
'    Dim delta_h() As Double ' deltah : 수위강하량
'
'    Dim radius() As Double ' 공반경
'    Dim Rw() As Double      ' 공반경 / 2000
'
'    Dim well_depth() As Double     ' 관정심도, well depth
'    Dim casing() As Double  ' 케이싱심도
'
'    Dim Q() As Double       '취수계획량
'    Dim delta_s() As Double
'    Dim hp() As Double
'
'    Dim daeSoo() As Double  ' 대수층 두께
'
'    Dim T1() As Double      ' T1
'    Dim T2() As Double      ' T2
'    Dim TA() As Double      ' TA - (T1+T2)/2, TAverage
'
'    Dim S1() As Double      ' S1
'    Dim S2() As Double      ' S2 - 스킨팩터 해석, s값
'
'    Dim T0() As Double         ' 단계양수시험의 T값
'    Dim S0() As Double         ' S값, 0.005: 피압대수층, 0.001: 누수대수층, 0.1: 자유면대수층
'    Dim ER_MODE() As String    ' 영향반경산정공식 선정
'    Dim ER1() As Double
'    Dim ER2() As Double
'    Dim ER3() As Double
'
'    Dim K() As Double
'    Dim time_() As Double   ' 안정수위도달시간
'
'    Dim shultze() As Double
'    Dim webber() As Double
'    Dim jacob() As Double
'
'    Dim skin() As Double ' skin factor
'    Dim er() As Double   ' effective radius, 유효우물반경
'
'
'    Dim qh() As Double ' 한계양수량
'    Dim qg() As Double ' 가채수량
'    Dim q1() As Double ' Q1
'
'
'    Dim sd1() As Double ' 1단계 수위강하량
'    Dim sd2() As Double ' 4단계 수위강하량
'
'
'    Dim C() As Double
'    Dim B() As Double
'
'    Dim ratio() As Double
'
'
' ' --------------------------------------------------------------------------------------
'
'    nofwell = GetNumberOfWell()
'    Sheets("YangSoo").Select
'
' ' --------------------------------------------------------------------------------------
'
'    ReDim natural(1 To nofwell)
'    ReDim stable(1 To nofwell)
'    ReDim recover(1 To nofwell)
'    ReDim delta_h(1 To nofwell)
'    ReDim Sw(1 To nofwell)
'
'
'    ReDim radius(1 To nofwell)
'    ReDim Rw(1 To nofwell)
'
'    ReDim well_depth(1 To nofwell)
'    ReDim casing(1 To nofwell)
'
'    ReDim Q(1 To nofwell)
'    ReDim delta_s(1 To nofwell)
'    ReDim hp(1 To nofwell)
'
'    ReDim daeSoo(1 To nofwell)
'
'    ReDim T1(1 To nofwell)
'    ReDim T2(1 To nofwell)
'    ReDim TA(1 To nofwell)
'
'    ReDim S1(1 To nofwell)
'    ReDim S2(1 To nofwell)
'
'    ReDim K(1 To nofwell)
'    ReDim time_(1 To nofwell)
'
'    ReDim shultze(1 To nofwell)
'    ReDim webber(1 To nofwell)
'    ReDim jacob(1 To nofwell)
'
'    ReDim skin(1 To nofwell)
'    ReDim er(1 To nofwell)
'
'    ReDim ER1(1 To nofwell)
'    ReDim ER2(1 To nofwell)
'    ReDim ER3(1 To nofwell)
'
'    ReDim qh(1 To nofwell)
'    ReDim qg(1 To nofwell)
'
'
'    ReDim sd1(1 To nofwell)
'    ReDim sd2(1 To nofwell)
'    ReDim q1(1 To nofwell)
'
'    ReDim C(1 To nofwell)
'    ReDim B(1 To nofwell)
'
'    ReDim ratio(1 To nofwell)
'
'
'    ReDim T0(1 To nofwell)         ' 단계양수시험의 T값
'    ReDim S0(1 To nofwell)         ' S값, 0.005: 피압대수층, 0.001: 누수대수층, 0.1: 자유면대수층
'    ReDim ER_MODE(1 To nofwell)
'
'
'
'    If Not (isSingleWellImport) And singleWell = 999 Then
'        rngString = "A5:AN" & (nofwell + 5 - 1)
'        Call EraseCellData(rngString)
'    End If
'
'    For i = 1 To nofwell
'
'        ' isSingleWellImport = True ---> SingleWell Import
'        ' isSingleWellImport = False ---> AllWell Import
'
'        If isSingleWellImport Then
'            If i = singleWell Then
'                GoTo SINGLE_ITERATION
'            Else
'                GoTo NEXT_ITERATION
'            End If
'        End If
'
'SINGLE_ITERATION:
'
'        fName = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
'        If Not IsWorkBookOpen(fName) Then
'            MsgBox "Please open the yangsoo data ! " & fName
'            Exit Sub
'        End If
'
'
'        Q(i) = Workbooks(fName).Worksheets("Input").Range("m51").value
'        hp(i) = Workbooks(fName).Worksheets("Input").Range("i48").value
'
'        natural(i) = Workbooks(fName).Worksheets("Input").Range("m48").value
'        stable(i) = Workbooks(fName).Worksheets("Input").Range("m49").value
'        radius(i) = Workbooks(fName).Worksheets("Input").Range("m44").value
'        Rw(i) = radius(i) / 2000
'
'        well_depth(i) = Workbooks(fName).Worksheets("Input").Range("m45").value
'        casing(i) = Workbooks(fName).Worksheets("Input").Range("i52").value
'
'
'        C(i) = Workbooks(fName).Worksheets("Input").Range("A31").value
'        B(i) = Workbooks(fName).Worksheets("Input").Range("B31").value
'
'
'
'        recover(i) = Workbooks(fName).Worksheets("SkinFactor").Range("c10").value
'        Sw(i) = stable(i) - recover(i)
'
'        delta_h(i) = Workbooks(fName).Worksheets("SkinFactor").Range("b16").value
'        delta_s(i) = Workbooks(fName).Worksheets("SkinFactor").Range("b4").value
'
'        daeSoo(i) = Workbooks(fName).Worksheets("SkinFactor").Range("c16").value
'
'        '----------------------------------------------------------------------------------
'
'        T0(i) = Workbooks(fName).Worksheets("SkinFactor").Range("d4").value
'        S0(i) = Workbooks(fName).Worksheets("SkinFactor").Range("f4").value
'        ER_MODE(i) = Workbooks(fName).Worksheets("SkinFactor").Range("h10").value
'
'        T1(i) = Workbooks(fName).Worksheets("SkinFactor").Range("d5").value
'        T2(i) = Workbooks(fName).Worksheets("SkinFactor").Range("h13").value
'        TA(i) = (T1(i) + T2(i)) / 2
'
'        S1(i) = Workbooks(fName).Worksheets("SkinFactor").Range("e10").value
'        S2(i) = Workbooks(fName).Worksheets("SkinFactor").Range("i16").value
'
'        K(i) = Workbooks(fName).Worksheets("SkinFactor").Range("e16").value
'        time_(i) = Workbooks(fName).Worksheets("SkinFactor").Range("h16").value
'
'        shultze(i) = Workbooks(fName).Worksheets("SkinFactor").Range("c13").value
'        webber(i) = Workbooks(fName).Worksheets("SkinFactor").Range("c18").value
'        jacob(i) = Workbooks(fName).Worksheets("SkinFactor").Range("c23").value
'
'        skin(i) = Workbooks(fName).Worksheets("SkinFactor").Range("g6").value
'        er(i) = Workbooks(fName).Worksheets("SkinFactor").Range("c8").value
'
'
'        ' 경험식, 1번, 2번, 3번의 유효우물반경
'        ER1(i) = Workbooks(fName).Worksheets("SkinFactor").Range("K8").value
'        ER2(i) = Workbooks(fName).Worksheets("SkinFactor").Range("K9").value
'        ER3(i) = Workbooks(fName).Worksheets("SkinFactor").Range("K10").value
'
'        '----------------------------------------------------------------------------------
'
'        qh(i) = Workbooks(fName).Worksheets("SafeYield").Range("b13").value
'        qg(i) = Workbooks(fName).Worksheets("SafeYield").Range("b7").value
'
'        sd1(i) = Workbooks(fName).Worksheets("SafeYield").Range("b3").value
'        sd2(i) = Workbooks(fName).Worksheets("SafeYield").Range("b4").value
'        q1(i) = Workbooks(fName).Worksheets("SafeYield").Range("b2").value
'
'        ratio(i) = Workbooks(fName).Worksheets("SafeYield").Range("b11").value
'
'        '*****************************************************************************************
'
'        Cells(4 + i, "a").value = "W-" & CStr(i)
'        Cells(4 + i, "b").value = natural(i)
'        Cells(4 + i, "c").value = stable(i)
'
'        Cells(4 + i, "d").value = recover(i)
'        Cells(4 + i, "d").NumberFormat = "0.00"
'
'        Cells(4 + i, "e").value = Sw(i)
'        Cells(4 + i, "e").NumberFormat = "0.00"
'
'        Cells(4 + i, "f").value = delta_h(i)
'        Cells(4 + i, "f").NumberFormat = "0.00"
'
'        Cells(4 + i, "g").value = radius(i)
'        Cells(4 + i, "h").value = Rw(i)
'        Cells(4 + i, "i").value = well_depth(i)
'        Cells(4 + i, "j").value = casing(i)
'        Cells(4 + i, "k").value = Q(i)
'
'        Cells(4 + i, "l").value = delta_s(i)
'        Cells(4 + i, "l").NumberFormat = "0.00"
'
'        Cells(4 + i, "m").value = hp(i)
'        Cells(4 + i, "n").value = daeSoo(i)
'
'        Cells(4 + i, "o").value = T1(i)
'        Cells(4 + i, "o").NumberFormat = "0.0000"
'
'        Cells(4 + i, "p").value = T2(i)
'        Cells(4 + i, "p").NumberFormat = "0.0000"
'
'        Cells(4 + i, "q").value = TA(i)
'        Cells(4 + i, "q").NumberFormat = "0.0000"
'
'        Cells(4 + i, "r").value = S1(i)
'
'        Cells(4 + i, "s").value = S2(i)
'        Cells(4 + i, "s").NumberFormat = "0.0000000"
'
'        Cells(4 + i, "t").value = K(i)
'        Cells(4 + i, "t").NumberFormat = "0.0000"
'
'        Cells(4 + i, "u").value = time_(i)
'
'        Cells(4 + i, "v").value = shultze(i)
'        Cells(4 + i, "v").NumberFormat = "0.0"
'
'        Cells(4 + i, "w").value = webber(i)
'        Cells(4 + i, "w").NumberFormat = "0.0"
'
'        Cells(4 + i, "x").value = jacob(i)
'        Cells(4 + i, "x").NumberFormat = "0.0"
'
'
'
'        Cells(4 + i, "y").value = Format(skin(i), "0.0000")
'
'        Cells(4 + i, "z").value = er(i)
'        Cells(4 + i, "z").NumberFormat = "0.0000"
'
'        Cells(4 + i, "aa").value = Format(qh(i), "0.")
'        Cells(4 + i, "ab").value = Format(qg(i), "0.00")
'        Cells(4 + i, "ac").value = Format(q1(i), "0.")
'
'        Cells(4 + i, "ad").value = Format(sd1(i), "0.00")
'        Cells(4 + i, "ae").value = Format(sd2(i), "0.00")
'
'        Cells(4 + i, "af").value = C(i)
'        Cells(4 + i, "ag").value = B(i)
'
'        Cells(4 + i, "ah").value = ratio(i)
'        Cells(4 + i, "ah").NumberFormat = "0.0%"
'
'
'        ' 2023/09/22 새로 추가한 값들 ...
'        Cells(4 + i, "AI").value = Format(T0(i), "0.0000")
'        Cells(4 + i, "AJ").value = Format(S0(i), "0.0000")
'        Cells(4 + i, "AK").value = ER_MODE(i)
'
'        Cells(4 + i, "AL").value = Format(ER1(i), "0.0000")
'        Cells(4 + i, "AM").value = Format(ER2(i), "0.0000")
'        Cells(4 + i, "AN").value = Format(ER3(i), "0.0000")
'
'NEXT_ITERATION:
'
'    Next i
'End Sub
