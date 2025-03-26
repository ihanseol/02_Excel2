Private Sub CommandButton1_Click()
    
    Sheets("Aggregate1").Visible = False
    Sheets("Well").Select
    
End Sub

'q1 - 한계양수량 - b13
'q2 - 가채수량 - b7
'q3 - 취수계획량 - b15
'ratio - b11
'qq1 - 1단계 양수량


' Agg1_Tentative_Water_Intake : 적정취수량의 계산
'
Private Sub CommandButton2_Click()
' Collect Data
    Call TurnOffStuff
    Call modAgg1.ImportAggregateData(999, False)
    Call TurnOnStuff
End Sub


' 영수시험 데이터 파일이름, 불러오기
Private Sub CommandButton3_Click()
    ' SingleWell Import
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = BaseData_ETC.GetOtherFileName
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
    
    Call modAgg1.ImportAggregateData(singleWell, True)

End Sub



