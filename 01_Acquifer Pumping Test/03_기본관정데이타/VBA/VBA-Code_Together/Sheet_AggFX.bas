
Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
  'Collect Data
    Dim fName As String
    
    fName = "A1_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "YangSoo File Does not OPEN ... ! " & fName
        Exit Sub
    End If
    
    Call GetBaseDataFromYangSoo(999, False)
End Sub


Private Sub CommandButton3_Click()
    ' Write Formula Button
       
       Call WriteFormula
    ' End of Write Formula Button
End Sub


Private Sub CommandButton4_Click()
    'single well import
    
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    ' 영수시험 데이터 파일이름, 불러오기
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
    
    Call GetBaseDataFromYangSoo(singleWell, True)

End Sub






