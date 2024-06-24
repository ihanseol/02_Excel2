Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim WBNAME As String
    Dim i As Long

    If Workbooks.count <> 2 Then
        GetOtherFileName = "NOTHING"
        Exit Function
    End If

    For Each Workbook In Application.Workbooks
        WBNAME = Workbook.name
        If StrComp(ThisWorkbook.name, WBNAME, vbTextCompare) = 0 Then
        ' �̸��� thisworkbook.name �� ���ٸ� , �����б��
            GoTo NEXT_ITERATION
        End If
        
        If ThisWorkbook.name <> Workbook.name And CheckSubstring(WBNAME, "����Ÿ") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    
    If ThisWorkbook.name <> WBNAME And CheckSubstring(WBNAME, "����Ÿ") Then
        GetOtherFileName = WBNAME
    Else
        GetOtherFileName = "NOTHING"
    End If
End Function


Function CheckSubstring(str As String, chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function

Sub CheckSheetExists(WB_NAME As String)
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.name = "All" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' Do something if sheet exists
    If sheetExists Then
        MsgBox "Sheet 'All' exists!"
        ' Place your code here to do something
    Else
        MsgBox "Sheet 'All' does not exist."
    End If
End Sub


Sub DuplicateWellSpec(ByVal this_WBNAME As String, ByVal well_no As Integer, obj As Class_Boolean)
    Dim WB_NAME As String
    Dim i As Integer
    Dim long_axis, short_axis, well_distance, well_height, surface_water_height As Long
    Dim degree_of_flow As Double


'    obj.Result = False, ��������
'    obj.Result = True , ��������
      
    If Workbooks.count <> 2 Then
        MsgBox "Please Open, �⺻��������Ÿ�� ����,  �⺻��������Ÿ ���� �ϳ��� �ҷ��ü��� �ֽ��ϴ�. ", vbOKOnly
        obj.Result = True
        Exit Sub
    End If
   
    
    WB_NAME = GetOtherFileName
    If WB_NAME = "NOTHING" Then
        GoTo SheetDoesNotExist
    End If
    
    On Error GoTo SheetDoesNotExist
    
    With Workbooks(WB_NAME).Worksheets(CStr(well_no))
        long_axis = .Range("K6").value
        short_axis = .Range("K7").value
        degree_of_flow = .Range("K12").value
        well_distance = .Range("K13").value
        well_height = .Range("K14").value
        surface_water_height = .Range("K15").value
    End With
    

    With Workbooks(this_WBNAME).Worksheets(CStr(well_no))
        .Range("K6") = long_axis
        .Range("K7") = short_axis
        .Range("K12") = degree_of_flow
        .Range("K13") = well_distance
        .Range("K14") = well_height
        .Range("K15") = surface_water_height
    End With
    
    obj.Result = False
    Exit Sub

SheetDoesNotExist:
    MsgBox "Please Open, �⺻��������Ÿ ������ �ƴմϴ�. ", vbOKOnly
    obj.Result = True
    
End Sub




Sub ImportWellSpec(ByVal well_no As Integer, obj As Class_Boolean)
    Dim WkbkName As Object
    Dim WBNAME As String
    Dim i As Integer
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir, skin As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, deltas As Double
    Dim casing As Integer

    WBNAME = "A" & GetNumeric2(well_no) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBNAME) Then
        MsgBox "Please open the yangsoo data ! " & WBNAME
        obj.Result = True
        Exit Sub
    Else
        obj.Result = False
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
    
    ' Skin Coefficient
    skin = Workbooks(WBNAME).Worksheets("SkinFactor").Range("G6").value
    
    ' yangsoo radius of influence
    RI1 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C13").value
    RI2 = Workbooks(WBNAME).Worksheets("SkinFactor").Range("C18").value
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
    
    Range("G4") = S1
    
    Range("h5") = skin 'skin coefficient
    Range("h6") = ir 'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(deltas, 2) 'deltas
End Sub
