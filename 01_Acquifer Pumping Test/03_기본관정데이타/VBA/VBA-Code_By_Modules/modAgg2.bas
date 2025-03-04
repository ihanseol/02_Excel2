Attribute VB_Name = "modAgg2"

' 회복 T값, S값 을 정리해서 뿌려준다.
Sub Write_SummaryTS(ByVal Well As Integer)

    Dim i As Integer

    i = Well - 1

    Range("H" & (i + 80)).value = "W-" & (i + 1)
    Range("i" & (i + 80)).value = Range("e" & (49 + i * 3)).value
    Range("J" & (i + 80)).value = Range("f" & (48 + i * 3)).value

End Sub

Sub ImportWellSpec(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim fName As String
    Dim nofwell, i As Integer

    Dim Q, Natural, Stable, Recover, Radius, DeltaS, DeltaH, DaeSoo, T1, T2, TA As Double
    Dim K, time_, S1, S2, Schultz, Webber, Jcob, Skin, Er As Double


    nofwell = GetNumberOfWell()
    Sheets("Aggregate2").Select

    Dim wsYangSoo As Worksheet
    Set wsYangSoo = Worksheets("YangSoo")


    If Not isSingleWellImport Then
    ' if All Colect Data Pressed ...

        'Write33
        Call EraseCellData("C3:J33")

        'Write34
        Call EraseCellData("L3:Q33")

        'Write35
        Call EraseCellData("S3:U33")

        'Write37
        Call EraseCellData("D37:AH43")

        'Write36
        Call EraseCellData("E48:F137")

        'Write38
        Call EraseCellData("H48:N77")

        'Write34
        Call EraseCellData("P48:S77")

        Call EraseCellData("H80:J109")

    End If


    For i = 1 To nofwell
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If

SINGLE_ITERATION:

        Q = wsYangSoo.Cells(4 + i, "k").value

        Natural = wsYangSoo.Cells(4 + i, "b").value
        Stable = wsYangSoo.Cells(4 + i, "c").value
        Recover = wsYangSoo.Cells(4 + i, "d").value

        Radius = wsYangSoo.Cells(4 + i, "h").value

        DeltaS = wsYangSoo.Cells(4 + i, "l").value
        DeltaH = wsYangSoo.Cells(4 + i, "f").value
        DaeSoo = wsYangSoo.Cells(4 + i, "n").value


        T1 = wsYangSoo.Cells(4 + i, "o").value
        T2 = wsYangSoo.Cells(4 + i, "p").value
        TA = wsYangSoo.Cells(4 + i, "q").value

        time_ = wsYangSoo.Cells(4 + i, "u").value

        S1 = wsYangSoo.Cells(4 + i, "r").value
        S2 = wsYangSoo.Cells(4 + i, "s").value
        K = wsYangSoo.Cells(4 + i, "t").value

        shultz = wsYangSoo.Cells(4 + i, "v").value
        Webber = wsYangSoo.Cells(4 + i, "w").value
        Jcob = wsYangSoo.Cells(4 + i, "x").value


        Skin = wsYangSoo.Cells(4 + i, "y").value
        Er = wsYangSoo.Cells(4 + i, "z").value

        Call TurnOffStuff

        Call modAgg2.WriteWellData_Single(Q, Natural, Stable, Recover, Radius, DeltaS, DaeSoo, T1, S1, i, isSingleWellImport)
        Call modAgg2.WriteData37_RadiusOfInfluence_Single(TA, K, S2, time_, DeltaH, DaeSoo, i, isSingleWellImport)
        Call modAgg2.WriteData36_TS_Analysis_Single(T1, T2, TA, S2, i, isSingleWellImport)
        Call modAgg2.Write38_RadiusOfInfluence_Result_Single(shultz, Webber, Jcob, i, isSingleWellImport)
        Call modAgg2.Wrote34_SkinFactor_Single(Skin, Er, i, isSingleWellImport)

        Call modAgg2.Write_SummaryTS(i)
        Call TurnOnStuff


NEXT_ITERATION:

    Next i

    Range("a1").Select
    Application.CutCopyMode = False

End Sub


' 3-3, 3-4, 3-5 결과출력
Sub WriteWellData_Single(Q As Variant, Natural As Variant, Stable As Variant, Recover As Variant, Radius As Variant, DeltaS As Variant, DaeSoo As Variant, T1 As Variant, S1 As Variant, ByVal i As Integer, ByVal isSingleWellImport As Boolean)

    Dim remainder As Integer

    ' 3-3, 장기양수시험결과 (Collect from yangsoo data)


    If isSingleWellImport Then
       EraseCellData ("C" & (i + 2) & ":J" & (i + 2))
       EraseCellData ("L" & (i + 2) & ":Q" & (i + 2))
       EraseCellData ("S" & (i + 2) & ":U" & (i + 2))
    End If

    Range("C" & (i + 2)).value = "W-" & i
    Range("D" & (i + 2)).value = 2880

    Range("e" & (i + 2)).value = Q
    Range("l" & (i + 2)).value = Q

    Range("f" & (i + 2)).value = Natural
    Range("g" & (i + 2)).value = Stable
    Range("h" & (i + 2)).value = Stable - Natural

    Range("i" & (i + 2)).value = Radius
    Range("j" & (i + 2)).value = DeltaS


    ' 3-4, aqtesolv 해석결과
    Range("m" & (i + 2)).value = Radius
    Range("n" & (i + 2)).value = Radius
    Range("o" & (i + 2)).value = DaeSoo
    Range("p" & (i + 2)).value = T1
    Range("q" & (i + 2)).value = S1


    '3-5, 수위회복시험 결과
    Range("s" & (i + 2)).value = Stable
    Range("t" & (i + 2)).value = Recover
    Range("u" & (i + 2)).value = Stable - Recover

    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(i + 2, "c"), Cells(i + 2, "j")), True)
            Call BackGroundFill(Range(Cells(i + 2, "l"), Cells(i + 2, "q")), True)
            Call BackGroundFill(Range(Cells(i + 2, "s"), Cells(i + 2, "u")), True)

    Else
            Call BackGroundFill(Range(Cells(i + 2, "c"), Cells(i + 2, "j")), False)
            Call BackGroundFill(Range(Cells(i + 2, "l"), Cells(i + 2, "q")), False)
            Call BackGroundFill(Range(Cells(i + 2, "s"), Cells(i + 2, "u")), False)
    End If

End Sub


' 3-7, 조사공별 수리상수
Sub WriteData37_RadiusOfInfluence_Single(TA As Variant, K As Variant, S2 As Variant, time_ As Variant, DeltaH As Variant, DaeSoo As Variant, i As Variant, ByVal isSingleWellImport As Boolean)

'****************************************
'    ip = 37 'W-1 point
'****************************************

    Dim ip, remainder As Variant
    Dim unit, rngString As String
    Dim Values As Variant

    Values = GetRowColumn("agg2_37_roi")
    ip = Values(2)


    If isSingleWellImport Then
        Call EraseCellData(ColumnNumberToLetter(4 + i) & ip & ":" & ColumnNumberToLetter(4 + i) & (ip + 6))
    End If

    Cells((ip + 0), (3 + i)).value = "W-" & i

    Cells((ip + 1), (3 + i)).value = TA
    Cells((ip + 1), (3 + i)).NumberFormat = "0.0000"

    Cells((ip + 2), (3 + i)).value = K
    Cells((ip + 2), (3 + i)).NumberFormat = "0.0000"


    Cells((ip + 3), (3 + i)).value = S2
    Cells((ip + 3), (3 + i)).NumberFormat = "0.0000000"

    Cells((ip + 4), (3 + i)).value = time_
    Cells((ip + 4), (3 + i)).NumberFormat = "0.0000"

    Cells((ip + 5), (3 + i)).value = DeltaH
    Cells((ip + 5), (3 + i)).NumberFormat = "0.00"

    Cells((ip + 6), (3 + i)).value = DaeSoo


    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + 1, (i + 3)), Cells(ip + 6, (i + 3))), True)
    Else
            Call BackGroundFill(Range(Cells(ip + 1, (i + 3)), Cells(ip + 6, (i + 3))), False)
    End If


End Sub


' 3-6, 수리상수산정결과
Sub WriteData36_TS_Analysis_Single(T1 As Variant, T2 As Variant, TA As Variant, S2 As Variant, i As Variant, ByVal isSingleWellImport As Boolean)

'****************************************
'    ip = 48
'****************************************
' Call EraseCellData("C48:F137")
' 137 - 48 = 89

    Dim ip, remainder As Variant
    Dim unit, rngString As String
    Dim Values As Variant
    Dim nofwell As Integer


    Values = GetRowColumn("agg2_36_surisangsoo")
    ip = Values(2)

    If isSingleWellImport Then
        Call EraseCellData("C" & (ip + (i - 1) * 3) & ":F" & (ip + (i - 1) * 3 + 2))
    End If

    Cells(ip + (i - 1) * 3, "C").value = "W-" & i

    Cells((ip + 0) + (i - 1) * 3, "D").value = "장기양수시험"
    Cells((ip + 1) + (i - 1) * 3, "D").value = "수위회복시험"
    Cells((ip + 2) + (i - 1) * 3, "D").value = "선택치"

    Cells((ip + 0) + (i - 1) * 3, "E").value = T1
    Cells((ip + 0) + (i - 1) * 3, "E").NumberFormat = "0.0000"

    Cells((ip + 1) + (i - 1) * 3, "E").value = T2
    Cells((ip + 1) + (i - 1) * 3, "E").NumberFormat = "0.0000"

    Cells((ip + 2) + (i - 1) * 3, "E").value = TA
    Cells((ip + 2) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    Cells((ip + 2) + (i - 1) * 3, "E").Font.Bold = True

    Cells((ip + 0) + (i - 1) * 3, "F").value = S2
    Cells((ip + 0) + ip + (i - 1) * 3, "F").NumberFormat = "0.0000000"

    Cells((ip + 2) + (i - 1) * 3, "F").value = S2
    Cells((ip + 2) + (i - 1) * 3, "F").NumberFormat = "0.0000000"
    Cells((ip + 2) + (i - 1) * 3, "F").Font.Bold = True


    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1) * 3, "C"), Cells((ip + 2) + (i - 1) * 3, "F")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1) * 3, "C"), Cells((ip + 2) + (i - 1) * 3, "F")), False)
    End If

End Sub



'3.8 영향반경
' 그리고 single 이 붙으면 알수있듯이 , 공번에 해당하는 라인에 관한것만 출력해준다.
'
Sub Write38_RadiusOfInfluence_Result_Single(shultz As Variant, Webber As Variant, Jcob As Variant, i As Variant, ByVal isSingleWellImport As Boolean)

'****************************************
'    ip = 48 'W-1 point
'****************************************
' Call EraseCellData("H48:N77")
' 77 - 48 = 29


    Dim ip, remainder As Variant
    Dim unit, rngString As String
    Dim Values As Variant

    Values = GetRowColumn("agg2_38_roi_result")
    ip = Values(2)

    'Call EraseCellData("H" & ip & ":N" & (ip + nofwell - 1))

    ' 단일공의 임포트의 경우 한줄만 지우고
    ' 그 이외에는 나머지 모두를 지워서 깨끗하게 해준다.

    If isSingleWellImport Then
        Call EraseCellData("H" & (ip + i - 1) & ":N" & (ip + i - 1))
    End If

    Cells(ip + (i - 1), "h").value = "W-" & i
    Cells(ip + (i - 1), "h").NumberFormat = "0.0"

    Cells(ip + (i - 1), "i").value = shultz
    Cells(ip + (i - 1), "i").NumberFormat = "0.0"

    Cells(ip + (i - 1), "j").value = Webber
    Cells(ip + (i - 1), "j").NumberFormat = "0.0"

    Cells(ip + (i - 1), "k").value = Jcob
    Cells(ip + (i - 1), "k").NumberFormat = "0.0"

    Cells(ip + (i - 1), "l").value = Round((shultz + Webber + Jcob) / 3, 1)
    Cells(ip + (i - 1), "l").NumberFormat = "0.0"

    Cells(ip + (i - 1), "m").value = Application.WorksheetFunction.max(shultz, Webber, Jcob)
    Cells(ip + (i - 1), "m").NumberFormat = "0.0"

    Cells(ip + (i - 1), "n").value = Application.WorksheetFunction.min(shultz, Webber, Jcob)
    Cells(ip + (i - 1), "n").NumberFormat = "0.0"


    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1), "h"), Cells(ip + (i - 1), "n")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1), "h"), Cells(ip + (i - 1), "n")), False)
    End If


End Sub



' 3.4 스킨계수
Sub Wrote34_SkinFactor_Single(Skin As Variant, Er As Variant, i As Variant, ByVal isSingleWellImport As Boolean)

'****************************************
'   ip = 48
'****************************************
' Call EraseCellData("P48:R77")
'****************************************

    Dim ip As Variant
    Dim unit, rngString As String
    Dim Values As Variant

    Values = GetRowColumn("agg2_34_skinfactor")
    ip = Values(2)


    If isSingleWellImport Then
        Call EraseCellData("P" & (ip + i - 1) & ":R" & (ip + i - 1))
    End If

    Cells(ip + (i - 1), "p").value = "W-" & i
    Cells(ip + (i - 1), "q").value = Skin
    Cells(ip + (i - 1), "q").NumberFormat = "0.0000"
    Cells(ip + (i - 1), "r").value = Er
    Cells(ip + (i - 1), "r").NumberFormat = "0.000"

    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1), "p"), Cells(ip + (i - 1), "r")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1), "p"), Cells(ip + (i - 1), "r")), False)
    End If

End Sub



