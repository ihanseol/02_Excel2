Option Explicit


' AggSum_Intake : �����ȹ��
' AggSum_Simdo : �����ɵ�
' AggSum_MotorHP : ��������
' AggSum_NaturalLevel : �ڿ�����
' AggSum_StableLevel : ��������
' AggSum_ToChool : ���ⱸ��
' AggSum_MotorSimdo : ���ͽɵ�
'

' AggSum_ROI : radius of influence
' AggSum_DI : drastic index
' AggSum_ROI_Stat :
' AggSum_26_AC : 26, AquiferCharacterization
' AggSum_26_RightAC : 26, Right AquiferCharacterizationn




Const WELL_BUFFER = 30


Private Sub EraseCellData(ByVal str_range As String)
    With Range(str_range)
        .Value = ""
    End With
End Sub

Private Sub CommandButton1_Click()
    Sheets("AggSum").Visible = False
    Sheets("Well").Select
End Sub


Private Sub Test_NameManager()
    Dim acColumn, acRow As Variant
    
    acColumn = Split(Range("ip_motor_simdo").Address, "$")(1)
    acRow = Split(Range("ip_motor_simdo").Address, "$")(2)
    
    '  Row = ActiveCell.Row
    '  col = ActiveCell.Column
    
    Debug.Print acColumn, acRow
End Sub

' Summary Button
Private Sub CommandButton2_Click()
    Dim nofwell As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggSum" Then Sheets("AggSum").Select


    ' Summary, Aquifer Characterization  Appropriated Water Analysis
    
    TurnOffStuff
    
    Call Write23_SummaryDevelopmentPotential
    Call Write26_AquiferCharacterization(nofwell)
    Call Write26_Right_AquiferCharacterization(nofwell)
    
    Call Write_RadiusOfInfluence(nofwell)
    Call Write_WaterIntake(nofwell)
    Call Check_DI
    
    Call Write_DiggingDepth(nofwell)
    Call Write_MotorPower(nofwell)
    Call Write_DrasticIndex(nofwell)
    
    Call Write_NaturalLevel(nofwell)
    Call Write_StableLevel(nofwell)
    
    
    Call Write_MotorTochool(nofwell)
    Call Write_MotorSimdo(nofwell)


    Range("D5").Select
    TurnOnStuff
    
End Sub

Sub Write23_SummaryDevelopmentPotential()
' Groundwater Development Potential, ���ϼ����߰��ɷ�
    
    Range("D4").Value = Worksheets(CStr(1)).Range("e17").Value
    Range("e4").Value = Worksheets(CStr(1)).Range("g14").Value
    Range("f4").Value = Worksheets(CStr(1)).Range("f19").Value
    Range("g4").Value = Worksheets(CStr(1)).Range("g13").Value
    
    Range("h4").Value = Worksheets(CStr(1)).Range("g19").Value
    Range("i4").Value = Worksheets(CStr(1)).Range("f21").Value
    Range("j4").Value = Worksheets(CStr(1)).Range("e21").Value
    Range("k4").Value = Worksheets(CStr(1)).Range("g21").Value
    
    ' --------------------------------------------------------------------
    
    Range("D8").Value = Worksheets(CStr(1)).Range("e17").Value
    Range("e8").Value = Worksheets(CStr(1)).Range("g14").Value
    Range("f8").Value = Worksheets(CStr(1)).Range("f19").Value
    Range("g8").Value = Worksheets(CStr(1)).Range("g13").Value
    
    Range("h8").Value = Worksheets(CStr(1)).Range("g19").Value
    Range("i8").Value = Worksheets(CStr(1)).Range("f21").Value
    Range("j8").Value = Worksheets(CStr(1)).Range("h19").Value
    Range("k8").Value = Worksheets(CStr(1)).Range("e21").Value

End Sub


'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>

Sub Write26_AquiferCharacterization(nofwell As Integer)
' ����� �м��� ����ä���� �м�
    Dim i, ip, remainder As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRowColumn("AggSum_26_AC")
    ip = Values(2)
    
    rngString = "D" & ip & ":J" & (ip + WELL_BUFFER - 1)
    Call EraseCellData(rngString)
        
    For i = 1 To nofwell
    
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(11 + i, "d"), Cells(11 + i, "j")), True)
        Else
                Call BackGroundFill(Range(Cells(11 + i, "d"), Cells(11 + i, "j")), False)
        End If
    
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "D").Value = "W-" & CStr(i)
        ' �ɵ�
        Cells(11 + i, "E").Value = Worksheets("Well").Cells(i + 3, 8).Value
        ' �����
        Cells(11 + i, "F").Value = Worksheets("Well").Cells(i + 3, 10).Value
        
        ' �ڿ�����
        Cells(11 + i, "G").Value = Worksheets(CStr(i)).Range("c20").Value
        Cells(11 + i, "G").NumberFormat = "0.00"
        
        ' ��������
        Cells(11 + i, "H").Value = Worksheets(CStr(i)).Range("c21").Value
        Cells(11 + i, "H").NumberFormat = "0.00"
        
        ' ���������
        Cells(11 + i, "I").Value = Worksheets(CStr(i)).Range("E7").Value
        Cells(11 + i, "I").NumberFormat = "0.0000"
        
        ' �������
        Cells(11 + i, "J").Value = Worksheets(CStr(i)).Range("G7").Value
        Cells(11 + i, "J").NumberFormat = "0.0000000"
    Next i
End Sub


Sub Write26_Right_AquiferCharacterization(nofwell As Integer)
' ����� �м��� ����ä���� �м�
    Dim i, ip, remainder As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRowColumn("AggSum_26_RightAC")
    ip = Values(2)
    
    rngString = "L" & ip & ":S" & (ip + WELL_BUFFER - 1)
    
    Call EraseCellData(rngString)
            
    For i = 1 To nofwell
    
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(11 + i, "L"), Cells(11 + i, "S")), True)
        Else
                Call BackGroundFill(Range(Cells(11 + i, "L"), Cells(11 + i, "S")), False)
        End If
    
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "L").Value = "W-" & CStr(i)
        ' �ɵ�
        Cells(11 + i, "M").Value = Worksheets("Well").Cells(i + 3, 8).Value
        ' �����
        Cells(11 + i, "N").Value = Worksheets("Well").Cells(i + 3, 10).Value
        
        ' �ڿ�����
        Cells(11 + i, "O").Value = Worksheets(CStr(i)).Range("c20").Value
        Cells(11 + i, "O").NumberFormat = "0.00"
        
        ' ��������
        Cells(11 + i, "P").Value = Worksheets(CStr(i)).Range("c21").Value
        Cells(11 + i, "P").NumberFormat = "0.00"
        
        '�������Ϸ�
        Cells(11 + i, "Q").Value = Worksheets(CStr(i)).Range("c21").Value - Worksheets(CStr(i)).Range("c20").Value
        Cells(11 + i, "Q").NumberFormat = "0.00"
        
        ' ���������
        Cells(11 + i, "R").Value = Worksheets(CStr(i)).Range("E7").Value
        Cells(11 + i, "R").NumberFormat = "0.0000"
         
        ' �������
        Cells(11 + i, "S").Value = Worksheets(CStr(i)).Range("G7").Value
        Cells(11 + i, "S").NumberFormat = "0.0000000"
    Next i
End Sub

'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>



Sub Write_RadiusOfInfluence(nofwell As Integer)
' �������ݰ�
    Dim i, ip, remainder As Integer
    Dim unit, rngString01, rngString02 As String
    Dim Values As Variant
        
    Values = GetRowColumn("AggSum_ROI")
    ip = Values(2)
    
    rngString01 = "D" & ip & ":G" & (ip + WELL_BUFFER - 1)
    rngString02 = "M" & ip & ":O" & (ip + WELL_BUFFER - 1)
    
    
    Call EraseCellData(rngString01)
    Call EraseCellData(rngString02)
        
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        ' WellNum
        Cells(ip - 1 + i, "D").Value = "W-" & CStr(i)
        ' �������ݰ�, �̰��� ������ ���� �ٸ���,
        ' �ϴ��� �ִ밪, shultz, webber, jcob �� �ִ밪�� �����ϴ°����� �Ѵ�.
        ' �׸��� �ʿ��� �κ���, �Ŀ� �߰������ش�.
        Cells(ip - 1 + i, "E").Value = Worksheets(CStr(i)).Range("H9").Value & unit
        Cells(ip - 1 + i, "F").Value = Worksheets(CStr(i)).Range("K6").Value & unit
        Cells(ip - 1 + i, "G").Value = Worksheets(CStr(i)).Range("K7").Value & unit
        
        
        '����ݰ��� �ִ�, �ּ�, ��հ��� �߰����ش�.
        Cells(ip - 1 + i, "M").Value = Worksheets(CStr(i)).Range("H9").Value & unit
        Cells(ip - 1 + i, "N").Value = Worksheets(CStr(i)).Range("H10").Value & unit
        Cells(ip - 1 + i, "O").Value = Worksheets(CStr(i)).Range("H11").Value & unit
        
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(ip - 1 + i, "d"), Cells(ip - 1 + i, "g")), True)
                Call BackGroundFill(Range(Cells(ip - 1 + i, "m"), Cells(ip - 1 + i, "o")), True)
        Else
                Call BackGroundFill(Range(Cells(ip - 1 + i, "d"), Cells(ip - 1 + i, "j")), False)
                Call BackGroundFill(Range(Cells(ip - 1 + i, "m"), Cells(ip - 1 + i, "o")), False)
        End If
        
        
    Next i
End Sub


Sub Write_DrasticIndex(nofwell As Integer)
' ���ƽ �ε���
    Dim i, ip, remainder As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRowColumn("AggSum_DI")
    ip = Values(2)
    
    rngString = "I" & Values(2) & ":K" & (Values(2) + WELL_BUFFER - 1)
    Call EraseCellData(rngString)
    
    For i = 1 To nofwell
        ' WellNum
        Cells(ip - 1 + i, "I").Value = "W-" & CStr(i)
        Cells(ip - 1 + i, "J").Value = Worksheets(CStr(i)).Range("k30").Value
        Cells(ip - 1 + i, "K").Value = Worksheets(CStr(i)).Range("k31").Value
        
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(ip - 1 + i, "i"), Cells(ip - 1 + i, "k")), True)
        Else
                Call BackGroundFill(Range(Cells(ip - 1 + i, "i"), Cells(ip - 1 + i, "k")), False)
        End If
        
    Next i
End Sub


'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>


Sub TestColumnLetter()

    ' ColumnNumberToLetter
    ' ColumnLetterToNumber
    
    Debug.Print ColumnLetterToNumber("D")
    Debug.Print ColumnLetterToNumber("AG")
    ' 4
    ' 33
    ' 33 = 4 + 30 - 1

End Sub

Sub Write_WaterIntake(nofwell As Integer)
' �����ȹ��
    Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_Intake")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    
    ' rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    rngString = StartCol & ip & ":AG79"
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = Sheets("drastic").Range("a16").Value
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        ' WellNum
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("C15").Value & unit
    Next i
End Sub


Sub Write_DiggingDepth(nofwell As Integer)
' �����ɵ�
   Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_Simdo")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    
    'rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    
    rngString = StartCol & ip & ":AG84"
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("c7").Value & unit
    Next i
End Sub



' Write_MotorTochool
' Write_MotorSimdo

Sub Write_MotorPower(nofwell As Integer)
' ���͸���
    Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_MotorHP")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    'rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    
    rngString = StartCol & ip & ":AG89"
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " Hp"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("c17").Value & unit
    Next i
End Sub



Sub Write_NaturalLevel(nofwell As Integer)
' �ڿ�����
    Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_NaturalLevel")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    ' rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    rngString = StartCol & ip & ":AG94"
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("c20").Value & unit
    Next i
End Sub

Sub Write_StableLevel(nofwell As Integer)
' ��������
    Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_StableLevel")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    'rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    
    rngString = StartCol & ip & ":AG99"
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("c21").Value & unit
    Next i
End Sub


Sub Write_MotorTochool(nofwell As Integer)
' ���ⱸ��
    Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_ToChool")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    'rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    rngString = StartCol & ip & ":AG104"
    
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " mm"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("c19").Value & unit
    Next i
End Sub



Sub Write_MotorSimdo(nofwell As Integer)
' ���ͽɵ�
    Dim i, ip As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    Dim StartCol, EndCol As String
    
    Values = GetRowColumn("AggSum_MotorSimdo")
    ip = Values(2)
    
    StartCol = Values(1)
    EndCol = ColumnNumberToLetter(ColumnLetterToNumber(StartCol) + nofwell - 1)
    
    
    ' rngString = StartCol & ip & ":" & EndCol & (ip + 1)
    
    rngString = StartCol & ip & ":AG109"
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.Value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(ip, (i + 3)).Value = "W-" & CStr(i)
        Cells(ip + 1, (i + 3)).Value = Worksheets(CStr(i)).Range("c18").Value & unit
    Next i
End Sub



Function CheckDrasticIndex(val As Integer) As String
    
    Dim Value As Integer
    Dim Result As String
    
    Select Case val
        Case Is <= 100
            Result = "�ſ쳷��"
        Case Is <= 120
            Result = "����"
        Case Is <= 140
            Result = "��������"
        Case Is <= 160
            Result = "�߰�����"
        Case Is <= 180
            Result = "����"
        Case Else
            Result = "�ſ����"
    End Select
    
    CheckDrasticIndex = Result
End Function


Sub Check_DI()
' Drastic Index �� ������ �߷��� ...

    Dim i, ip_Row, ip_Column As Integer
    Dim unit, rngString01 As String
    Dim Values As Variant
    
    Values = GetRowColumn("AggSum_Statistic_DrasticIndex")
    
    ip_Column = ColumnLetterToNumber(Values(1))
    ip_Row = Values(2)
    
    Range(ColumnNumberToLetter(ip_Column + 1) & ip_Row).Value = CheckDrasticIndex(Range(ColumnNumberToLetter(ip_Column) & ip_Row))
    Range(ColumnNumberToLetter(ip_Column + 1) & (ip_Row + 1)).Value = CheckDrasticIndex(Range(ColumnNumberToLetter(ip_Column) & (ip_Row + 1)))

End Sub
