Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub

'
' 2023/3/15, make summary page
' i think, this procedure made before .. but ican't the source excel file ..
' so i make again
' get page from pagenum
' quotient, remainder
' pagenum - 7 : 7/3 - 2, 1 (if remainder 1 = w1, remain = 2, w2)
'
Private Sub CommandButton3_Click()
' make summary page

    Dim n_sheets As Integer
    Dim result() As Integer
    Dim w3page, wselect, restpage As Integer
    'wselect = 1 --> only w1
    'wselect = 2 --> w1, w2
    
    
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
    n_sheets = sheets_count()
    result = DivideWellsBy3(n_sheets)
    
    
    ' result(0) = quotient
    ' result(1) = remainder
    w3page = result(0)
    
    Select Case result(1)
        Case 0
            restpage = 0
            wselect = 0
            
        Case 1
            restpage = 1
            wselect = 1
            
        Case 2
            restpage = 1
            wselect = 2
    End Select
    
    
    Call DuplicateQ3Page(w3page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call modWaterQualityTest.DuplicateRestQ3(wselect, w3page)
    End If
    
End Sub

'
' 2024/2/28, Get Water Spec From YangSoo
'

Private Sub CommandButton2_Click()

  Call GetWaterSpecFromYangSoo_Q3

End Sub



Private Sub CommandButton4_Click()
 
 Call modWaterQualityTest.DeleteAllSummaryPage("Q3")
   
End Sub

Private Sub CommandButton5_Click()
' get ec, ph, temp

    Call DataAnalysis
End Sub


' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub DataAnalysis()
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
        lowEC(i) = getEC_Q3(cellLOW, i)
        hiEC(i) = getEC_Q3(cellHI, i)
        
        lowPH(i) = getPH_Q3(cellLOW, i)
        hiPH(i) = getPH_Q3(cellHI, i)
        
        lowTEMP(i) = getTEMP_Q3(cellLOW, i)
        hiTEMP(i) = getTEMP_Q3(cellHI, i)
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


