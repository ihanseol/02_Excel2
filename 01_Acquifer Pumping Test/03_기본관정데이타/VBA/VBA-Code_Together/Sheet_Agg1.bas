Private Sub CommandButton1_Click()
    
    Sheets("Aggregate1").Visible = False
    Sheets("Well").Select
    
End Sub

'q1 - �Ѱ����� - b13
'q2 - ��ä���� - b7
'q3 - �����ȹ�� - b15
'ratio - b11
'qq1 - 1�ܰ� �����


' Agg1_Tentative_Water_Intake : ����������� ���



Private Sub CommandButton2_Click()
' Collect Data

Call AggregateOne_Import(999, False)

End Sub



Private Sub CommandButton3_Click()
    ' SingleWell Import
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = GetOtherFileName
    'MsgBox WB_NAME
    
    'If Workbook Is Nothing Then
    '    GetOtherFileName = "Empty"
    'Else
    '    GetOtherFileName = Workbook.name
    'End If
        
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call AggregateOne_Import(singleWell, True)

End Sub



