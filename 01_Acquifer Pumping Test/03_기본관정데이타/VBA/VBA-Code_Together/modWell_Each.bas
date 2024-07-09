
Sub ImporEachWell(ByVal well_no As Integer)
'
' well_no -- well number
'
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, skin As Double
    
    ' nl : natural level, sl : stable level
    ' s3 - Recover Test �� S��
    
    Dim nl, sl, deltas As Double
    Dim casing As Integer
    Dim wsYangSoo As Worksheet
    
    i = well_no
    Set wsYangSoo = Worksheets("YangSoo")
    BaseData_ETC_02.TurnOffStuff
    
    ' delta s : ����1���� ��������
    deltas = wsYangSoo.Cells(4 + i, "L").value
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
    nl = wsYangSoo.Cells(4 + i, "B").value
    sl = wsYangSoo.Cells(4 + i, "C").value
    casing = wsYangSoo.Cells(4 + i, "J").value
    
    
    T1 = wsYangSoo.Cells(4 + i, "O").value
    S1 = wsYangSoo.Cells(4 + i, "R").value
    T2 = wsYangSoo.Cells(4 + i, "P").value
    S2 = wsYangSoo.Cells(4 + i, "S").value
    S3 = wsYangSoo.Cells(4 + i, "AQ").value
    
    ' ��Ų���
    skin = wsYangSoo.Cells(4 + i, "Y").value
    
    ' yangsoo radius of influence
    RI1 = wsYangSoo.Cells(4 + i, "V").value  ' schultze
    RI2 = wsYangSoo.Cells(4 + i, "W").value  ' webber
    RI3 = wsYangSoo.Cells(4 + i, "X").value  ' jcob
    
    ' ��ȿ�칰�ݰ� , �������� ����
    ' ir = GetEffectiveRadius(WBNAME)
    ir = GetEffectiveRadiusFromFX(i)
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
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


Private Sub ImportEachWell_OLD()
    Dim WkbkName As Object
    Dim WBNAME, cell1 As String
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, skin As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, deltas As Double
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

    ' delta s : ����1���� ��������
    deltas = Workbooks(WBNAME).Worksheets("SkinFactor").Range("b4").value
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
    nl = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i4").value
    sl = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i6").value
    casing = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i10").value
    
    ' WkbkName.Close
    T1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("D5").value
    S1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("E10").value
    T2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("H13").value
    S2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i16").value
    S3 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("i13").value
    
    skin = Workbooks(WBNAME).Worksheets("SkinFactor").Range("G6").value
    
    ' yangsoo radius of influence
    ' ����, ����ݰ�
    RI1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C13").value
    ' ����, ����ݰ�
    RI2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C18").value
    ' ������, ����ݰ�
    RI3 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C23").value
    
    ' ��ȿ�칰�ݰ� , �������� ����
    ir = GetEffectiveRadius(WBNAME)
    
    ' �ڿ�����, ��������, ���̽� �ɵ� ����
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
