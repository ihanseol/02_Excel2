Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub



Private Sub CommandButton5_Click()
' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...

    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    
 
    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC_Q2(cellLOW, i)
        hiEC(i) = getEC_Q2(cellHI, i)
        
        lowPH(i) = getPH_Q2(cellLOW, i)
        hiPH(i) = getPH_Q2(cellHI, i)
        
        lowTEMP(i) = getTEMP_Q2(cellLOW, i)
        hiTEMP(i) = getTEMP_Q2(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"

End Sub



Private Sub CommandButton3_Click()
' make summary page

    Dim result() As Integer
    Dim w2page, wselect, restpage As Integer
    'wselect = 1 --> only w1
       
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
       
       
    result = DivideWellsBy2(sheets_count())
    
    ' result(0) = quotient
    ' result(1) = remainder
    
    w2page = result(0)
    restpage = result(1)
    
    Call DuplicateQ2Page(w2page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call modWaterQualityTest.DuplicateRestQ2(w2page)
    End If

End Sub


Private Sub CommandButton2_Click()
' get waterspec from yangsoo
  
  Call GetWaterSpecFromYangSoo_Q2

  
End Sub



Private Sub CommandButton4_Click()

 Call modWaterQualityTest.DeleteAllSummaryPage("Q2")

End Sub

