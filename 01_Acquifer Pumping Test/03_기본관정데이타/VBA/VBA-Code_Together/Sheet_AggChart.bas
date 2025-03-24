Option Explicit
'Sheet_AggChart


Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("AggChart").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
'Collect Data

   Call TurnOffStuff
   Call WriteAllCharts(999, False)
   Call TurnOnStuff

End Sub


Private Sub CommandButton3_Click()
    'single well import
    
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    'If Workbook Is Nothing Then
    '    GetOtherFileName = "Empty"
    'Else
    '    GetOtherFileName = Workbook.name
    'End If
            
    
    ' 영수시험 데이터 파일이름, 불러오기
    WB_NAME = BaseData_ETC.GetOtherFileName
    
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call WriteAllCharts(singleWell, True)

End Sub




