Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    ' Collect All Data
    Call modAgg2.GROK_ImportWellSpec(999, False)
End Sub




' 영수시험 데이터 파일이름, 불러오기
Private Sub CommandButton3_Click()
    ' SingleWell Import
    ' 지열공 같은경우, 단일공만 임포트 해야 할경우에 ....
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = BaseData_ETC.GetOtherFileName
    'MsgBox WB_NAME
    
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call modAgg2.ImportWellSpec(singleWell, True)

End Sub







