Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .Value = ""
    End With
End Sub


Private Sub CommandButton2_Click()
    ' Collect All Data

    Call ImportWellSpec(999, False)
End Sub


Private Sub CommandButton3_Click()
    ' SingleWell Import
    ' ������ �������, ���ϰ��� ����Ʈ �ؾ� �Ұ�쿡 ....
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = GetOtherFileName
    'MsgBox WB_NAME
    
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call ImportWellSpec(singleWell, True)

End Sub



' ȸ�� T��, S�� �� �����ؼ� �ѷ��ش�.
Private Sub Write_SummaryTS(ByVal Well As Integer)

    Dim i As Integer
    
    i = Well - 1
    
    Range("H" & (i + 80)).Value = "W-" & (i + 1)
    Range("i" & (i + 80)).Value = Range("e" & (49 + i * 3)).Value
    Range("J" & (i + 80)).Value = Range("f" & (48 + i * 3)).Value

End Sub




Private Sub ImportWellSpec(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Q, natural, stable, recover, radius, deltas, deltah, daeSoo, T1, T2, TA As Double
    Dim K, time_, S1, S2, schultz, webber, jcob, skin, er As Double
    

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
        Call EraseCellData("E37:AH43")
        
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
   
        Q = wsYangSoo.Cells(4 + i, "k").Value
        
        natural = wsYangSoo.Cells(4 + i, "b").Value
        stable = wsYangSoo.Cells(4 + i, "c").Value
        recover = wsYangSoo.Cells(4 + i, "d").Value
        
        radius = wsYangSoo.Cells(4 + i, "h").Value
        
        deltas = wsYangSoo.Cells(4 + i, "l").Value
        deltah = wsYangSoo.Cells(4 + i, "f").Value
        daeSoo = wsYangSoo.Cells(4 + i, "n").Value
        
        
        T1 = wsYangSoo.Cells(4 + i, "o").Value
        T2 = wsYangSoo.Cells(4 + i, "p").Value
        TA = wsYangSoo.Cells(4 + i, "q").Value
        
        time_ = wsYangSoo.Cells(4 + i, "u").Value
                
        S1 = wsYangSoo.Cells(4 + i, "r").Value
        S2 = wsYangSoo.Cells(4 + i, "s").Value
        K = wsYangSoo.Cells(4 + i, "t").Value
        
        shultz = wsYangSoo.Cells(4 + i, "v").Value
        webber = wsYangSoo.Cells(4 + i, "w").Value
        jcob = wsYangSoo.Cells(4 + i, "x").Value
        
        
        skin = wsYangSoo.Cells(4 + i, "y").Value
        er = wsYangSoo.Cells(4 + i, "z").Value
        
        Call TurnOffStuff
        
        Call WriteWellData_Single(Q, natural, stable, recover, radius, deltas, daeSoo, T1, S1, i, isSingleWellImport)
        Call WriteData37_RadiusOfInfluence_Single(TA, K, S2, time_, deltah, daeSoo, i, isSingleWellImport)
        Call WriteData36_TS_Analysis_Single(T1, T2, TA, S2, i, isSingleWellImport)
        Call Write38_RadiusOfInfluence_Result_Single(shultz, webber, jcob, i, isSingleWellImport)
        Call Wrote34_SkinFactor_Single(skin, er, i, isSingleWellImport)
        
        Call Write_SummaryTS(i)
        
        Call TurnOnStuff
        
    
NEXT_ITERATION:
    
    Next i

    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub


' 3-3, 3-4, 3-5 ������
Sub WriteWellData_Single(Q As Variant, natural As Variant, stable As Variant, recover As Variant, radius As Variant, deltas As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, ByVal i As Integer, ByVal isSingleWellImport As Boolean)
    
    Dim remainder As Integer
    
    ' 3-3, ����������� (Collect from yangsoo data)
    
    
    If isSingleWellImport Then
       EraseCellData ("C" & (i + 2) & ":J" & (i + 2))
       EraseCellData ("L" & (i + 2) & ":Q" & (i + 2))
       EraseCellData ("S" & (i + 2) & ":U" & (i + 2))
    End If
    
    Range("C" & (i + 2)).Value = "W-" & i
    Range("D" & (i + 2)).Value = 2880
    
    Range("e" & (i + 2)).Value = Q
    Range("l" & (i + 2)).Value = Q
    
    Range("f" & (i + 2)).Value = natural
    Range("g" & (i + 2)).Value = stable
    Range("h" & (i + 2)).Value = stable - natural
    
    Range("i" & (i + 2)).Value = radius
    Range("j" & (i + 2)).Value = deltas
    
    
    ' 3-4, aqtesolv �ؼ����
    Range("m" & (i + 2)).Value = radius
    Range("n" & (i + 2)).Value = radius
    Range("o" & (i + 2)).Value = daeSoo
    Range("p" & (i + 2)).Value = T1
    Range("q" & (i + 2)).Value = S1
    
    
    '3-5, ����ȸ������ ���
    Range("s" & (i + 2)).Value = stable
    Range("t" & (i + 2)).Value = recover
    Range("u" & (i + 2)).Value = stable - recover
    
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


' 3-7, ������� �������
Sub WriteData37_RadiusOfInfluence_Single(TA As Variant, K As Variant, S2 As Variant, time_ As Variant, deltah As Variant, daeSoo As Variant, i As Variant, ByVal isSingleWellImport As Boolean)

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
    
    Cells((ip + 0), (4 + i)).Value = "W-" & i
    
    Cells((ip + 1), (4 + i)).Value = TA
    Cells((ip + 1), (4 + i)).NumberFormat = "0.0000"
    
    Cells((ip + 2), (4 + i)).Value = K
    Cells((ip + 2), (4 + i)).NumberFormat = "0.0000"
    
    
    Cells((ip + 3), (4 + i)).Value = S2
    Cells((ip + 3), (4 + i)).NumberFormat = "0.0000000"
    
    Cells((ip + 4), (4 + i)).Value = time_
    Cells((ip + 4), (4 + i)).NumberFormat = "0.0000"
    
    Cells((ip + 5), (4 + i)).Value = deltah
    Cells((ip + 5), (4 + i)).NumberFormat = "0.00"
    
    Cells((ip + 6), (4 + i)).Value = daeSoo
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + 1, (i + 4)), Cells(ip + 6, (i + 4))), True)
    Else
            Call BackGroundFill(Range(Cells(ip + 1, (i + 4)), Cells(ip + 6, (i + 4))), False)
    End If
    

End Sub




' 3-6, ��������������
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
    
    Cells(ip + (i - 1) * 3, "C").Value = "W-" & i
            
    Cells((ip + 0) + (i - 1) * 3, "D").Value = "���������"
    Cells((ip + 1) + (i - 1) * 3, "D").Value = "����ȸ������"
    Cells((ip + 2) + (i - 1) * 3, "D").Value = "����ġ"

    Cells((ip + 0) + (i - 1) * 3, "E").Value = T1
    Cells((ip + 0) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    
    Cells((ip + 1) + (i - 1) * 3, "E").Value = T2
    Cells((ip + 1) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    
    Cells((ip + 2) + (i - 1) * 3, "E").Value = TA
    Cells((ip + 2) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    Cells((ip + 2) + (i - 1) * 3, "E").Font.Bold = True
    
    Cells((ip + 0) + (i - 1) * 3, "F").Value = S2
    Cells((ip + 0) + ip + (i - 1) * 3, "F").NumberFormat = "0.0000000"
    
    Cells((ip + 2) + (i - 1) * 3, "F").Value = S2
    Cells((ip + 2) + (i - 1) * 3, "F").NumberFormat = "0.0000000"
    Cells((ip + 2) + (i - 1) * 3, "F").Font.Bold = True
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1) * 3, "C"), Cells((ip + 2) + (i - 1) * 3, "F")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1) * 3, "C"), Cells((ip + 2) + (i - 1) * 3, "F")), False)
    End If

End Sub



'3.8 ����ݰ�
' �׸��� single �� ������ �˼��ֵ��� , ������ �ش��ϴ� ���ο� ���Ѱ͸� ������ش�.
'
Sub Write38_RadiusOfInfluence_Result_Single(shultz As Variant, webber As Variant, jcob As Variant, i As Variant, ByVal isSingleWellImport As Boolean)
 
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
    
    ' ���ϰ��� ����Ʈ�� ��� ���ٸ� �����
    ' �� �̿ܿ��� ������ ��θ� ������ �����ϰ� ���ش�.
    
    If isSingleWellImport Then
        Call EraseCellData("H" & (ip + i - 1) & ":N" & (ip + i - 1))
    End If
    
    Cells(ip + (i - 1), "h").Value = "W-" & i
    Cells(ip + (i - 1), "h").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "i").Value = shultz
    Cells(ip + (i - 1), "i").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "j").Value = webber
    Cells(ip + (i - 1), "j").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "k").Value = jcob
    Cells(ip + (i - 1), "k").NumberFormat = "0.0"

    Cells(ip + (i - 1), "l").Value = Round((shultz + webber + jcob) / 3, 1)
    Cells(ip + (i - 1), "l").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "m").Value = Application.WorksheetFunction.max(shultz, webber, jcob)
    Cells(ip + (i - 1), "m").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "n").Value = Application.WorksheetFunction.min(shultz, webber, jcob)
    Cells(ip + (i - 1), "n").NumberFormat = "0.0"
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1), "h"), Cells(ip + (i - 1), "n")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1), "h"), Cells(ip + (i - 1), "n")), False)
    End If


End Sub



' 3.4 ��Ų���
Sub Wrote34_SkinFactor_Single(skin As Variant, er As Variant, i As Variant, ByVal isSingleWellImport As Boolean)
    
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
    
    Cells(ip + (i - 1), "p").Value = "W-" & i
    Cells(ip + (i - 1), "q").Value = skin
    Cells(ip + (i - 1), "q").NumberFormat = "0.0000"
    Cells(ip + (i - 1), "r").Value = er
    Cells(ip + (i - 1), "r").NumberFormat = "0.000"
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1), "p"), Cells(ip + (i - 1), "r")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1), "p"), Cells(ip + (i - 1), "r")), False)
    End If

End Sub







