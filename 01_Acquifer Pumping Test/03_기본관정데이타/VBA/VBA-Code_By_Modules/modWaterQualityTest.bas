Attribute VB_Name = "modWaterQualityTest"

Sub DeleteAllSummaryPage(ByVal well_str As String)
' delete all summary page

    Dim nof_p, i As Integer
    nof_p = GetNumberOf_P
    
    For i = 1 To nof_p
        Application.DisplayAlerts = False
        On Error Resume Next
        
        Worksheets("p" & i).Delete
        
        On Error GoTo 0
        Application.DisplayAlerts = True
    Next i
    
    Sheets(well_str).Activate
End Sub



Sub GetWaterSpecFromYangSoo_Q1()
  Dim thisname, fName As String
  Dim cell  As String
  Dim Time As Date
  Dim bTemp, ec1, ph1 As Double
  
  
  cell = Range("d12").value
  
  thisname = ActiveWorkbook.name
  fName = "A" & GetNumeric2(cell) & "_ge_OriginalSaveFile.xlsm"
 
  If Not IsWorkBookOpen(fName) Then
    MsgBox "Please open the yangsoo data ! " & fName
    Exit Sub
  End If
  
  ' Range("k2") = fname
   
  '------------------------------------------------------------------------
  Time = Workbooks(fName).Worksheets("w1").Range("c6").value
  bTemp = Workbooks(fName).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fName).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fName).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  Range("c6").value = Time
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Call TurnOffStuff
  Call getModDataFromYangSooSingle(thisname, fName)
  Call TurnOnStuff
End Sub


Sub GetWaterSpecFromYangSoo_Q2()
  Dim thisname, fname1, fname2 As String
  Dim cell1, cell2 As String
  Dim time1 As Date
  Dim bTemp1, bTemp2, ec1, ec2, ph1, ph2 As Double
  
  
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  
  thisname = ActiveWorkbook.name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  ' Range("k2") = fname1
  ' Range("k3") = fname2
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  bTemp1 = Workbooks(fname1).Worksheets("w1").Range("c7").value
  bTemp2 = Workbooks(fname2).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp1
  Range("d7").value = bTemp2
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  Call TurnOffStuff
  Call getModDataFromYangSooDual(thisname, fname1)
  Call getModDataFromYangSooDual(thisname, fname2)
  Call TurnOnStuff
End Sub



Sub GetWaterSpecFromYangSoo_Q3()
  Dim thisname, fname1, fname2, fname3 As String
  Dim cell1, cell2, cell3 As String
  Dim time1 As Date
  Dim bTemp, bTemp2, bTemp3, ec1, ec2, ec3, ph1, ph2, ph3 As Double
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  cell3 = Range("j12").value
  
  
  thisname = ActiveWorkbook.name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  fname3 = "A" & GetNumeric2(cell3) & "_ge_OriginalSaveFile.xlsm"
   
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname3) Then
    MsgBox "Please open the yangsoo data ! " & fname3
    Exit Sub
  End If
  
  'Range("k2") = fname1
  'Range("k3") = fname2
  'Range("k4") = fname3
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  
  bTemp = Workbooks(fname1).Worksheets("w1").Range("c7").value
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  
  
  bTemp2 = Workbooks(fname2).Worksheets("w1").Range("c7").value
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  bTemp3 = Workbooks(fname3).Worksheets("w1").Range("c7").value
  ec3 = Workbooks(fname3).Worksheets("w1").Range("c8").value
  ph3 = Workbooks(fname3).Worksheets("w1").Range("c9").value
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
  Range("d7").value = bTemp2
  Range("e7").value = bTemp3
  
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  Range("e8").value = ec3
  Range("e9").value = ph3
  
  Call TurnOffStuff
  Call getModDataFromYangSooTripple(thisname, fname1)
  Call getModDataFromYangSooTripple(thisname, fname2)
  Call getModDataFromYangSooTripple(thisname, fname3)
  Call TurnOnStuff

End Sub




'******************************************************************************************************************************




Sub getModDataFromYangSooSingle(ByVal thisname As String, ByVal fName As String)
    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    Range("h14").Select
   
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub


Sub getModDataFromYangSooDual(ByVal thisname As String, ByVal fName As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fName)) Mod 2

    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("h31").Select
    Else
        Range("d31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub



Sub getModDataFromYangSooTripple(ByVal thisname As String, ByVal fName As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fName)) Mod 3

    Windows(fName).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("l31").Select
    ElseIf f = 1 Then
        Range("d31").Select
    Else
        Range("h31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub


'******************************************************************************************************************************



Function getEC_Q1(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & well).Activate
    
    If LOWHI = cellLOW Then
        getEC_Q1 = Sheets("p" & CStr(well)).Range("e25").value
    Else
        getEC_Q1 = Sheets("p" & CStr(well)).Range("e24").value
    End If
End Function

Function getPH_Q1(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate
    
    If LOWHI = cellLOW Then
        getPH_Q1 = Sheets("p" & CStr(well)).Range("f25").value
    Else
        getPH_Q1 = Sheets("p" & CStr(well)).Range("f24").value
    End If
    
End Function

Function getTEMP_Q1(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate

    If LOWHI = cellLOW Then
        getTEMP_Q1 = Sheets("p" & CStr(well)).Range("d25").value
    Else
        getTEMP_Q1 = Sheets("p" & CStr(well)).Range("d24").value
    End If
End Function


Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    
    Application.DisplayAlerts = False
    IsSheet = True
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function



Sub DuplicateQ1Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & n
    
    With ActiveSheet.Tab
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
    
    Call SetWellPropertyQ1(n)
    
End Sub

Sub SetWellPropertyQ1(ByVal i As Integer)
    ActiveSheet.Range("C4") = "W-" & CStr(i)
    ActiveSheet.Range("D12") = "W-" & CStr(i)
    ActiveSheet.Range("H12") = "W-" & CStr(i)
    
    
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub


'******************************************************************************************************************************


Function getEC_Q2(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getEC_Q2 = Sheets("p" & CStr(page)).Range("e25").value
        Else
            getEC_Q2 = Sheets("p" & CStr(page)).Range("e24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getEC_Q2 = Sheets("p" & CStr(page)).Range("h25").value
        Else
            getEC_Q2 = Sheets("p" & CStr(page)).Range("h24").value
        End If
    End If
End Function

Function getPH_Q2(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getPH_Q2 = Sheets("p" & CStr(page)).Range("f25").value
        Else
            getPH_Q2 = Sheets("p" & CStr(page)).Range("f24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getPH_Q2 = Sheets("p" & CStr(page)).Range("i25").value
        Else
            getPH_Q2 = Sheets("p" & CStr(page)).Range("i24").value
        End If
    End If
End Function

Function getTEMP_Q2(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("d25").value
        Else
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("d24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("g25").value
        Else
            getTEMP_Q2 = Sheets("p" & CStr(page)).Range("g24").value
        End If
    End If
End Function


Sub DuplicateQ2Page(ByVal n As Integer)
' n : Q2 page 복사할 회수
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .themeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ2(i)
    Next i
End Sub


Sub SetWellPropertyQ2(ByVal i As Integer)
' i : index of well

    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 2 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 2 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 2 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 2 + 2)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
End Sub


Sub SetWellPropertyRestQ2(ByVal w2page As Integer)
    Dim firstwell As Integer
      
    firstwell = 2 * w2page + 1
    
    ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
    ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub



Function DivideWellsBy2(ByVal numberOfWells As Integer) As Integer()
    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = (numberOfWells - 1) \ 2
    remainder = numberOfWells Mod 2
    
    
    If remainder = 0 Then
        result(0) = quotient + 1
    Else
        result(0) = quotient
    End If
    
    result(1) = remainder
    
    DivideWellsBy2 = result
End Function

Sub DuplicateRestQ2(ByVal w2page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & CStr(w2page + 1)
    
    With ActiveSheet.Tab
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRestQ2(w2page)
    
End Sub


'**********************************************************************************************************




' 1, 2, 3 --> p1
' 4, 5, 6 --> p2

Function getEC_Q3(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getEC_Q3 = Sheets("p" & CStr(page)).Range("k25").value
            Else
                getEC_Q3 = Sheets("p" & CStr(page)).Range("k24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getEC_Q3 = Sheets("p" & CStr(page)).Range("e25").value
            Else
                getEC_Q3 = Sheets("p" & CStr(page)).Range("e24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getEC_Q3 = Sheets("p" & CStr(page)).Range("h25").value
            Else
                getEC_Q3 = Sheets("p" & CStr(page)).Range("h24").value
            End If
    End Select
End Function

Function getPH_Q3(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getPH_Q3 = Sheets("p" & CStr(page)).Range("l25").value
            Else
                getPH_Q3 = Sheets("p" & CStr(page)).Range("l24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getPH_Q3 = Sheets("p" & CStr(page)).Range("f25").value
            Else
                getPH_Q3 = Sheets("p" & CStr(page)).Range("f24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getPH_Q3 = Sheets("p" & CStr(page)).Range("i25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("i24").value
            End If
    End Select
End Function

Function getTEMP_Q3(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("J25").value
            Else
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("J24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("d25").value
            Else
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("d24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("g25").value
            Else
                getTEMP_Q3 = Sheets("p" & CStr(page)).Range("g24").value
            End If
    End Select
End Function




Function DivideWellsBy3(ByVal numberOfWells As Integer) As Integer()

    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = numberOfWells \ 3
    remainder = numberOfWells Mod 3
    
    result(0) = quotient
    result(1) = remainder
    
    DivideWellsBy3 = result
    
End Function




Sub SetWellPropertyQ3(ByVal i As Integer)
    
    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 3 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 3 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 3 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 3 + 2)
    
    ActiveSheet.Range("J12") = "W-" & CStr((i - 1) * 3 + 3)
    ActiveSheet.Range("L29") = "W-" & CStr((i - 1) * 3 + 3)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    
End Sub

Sub SetWellPropertyRestQ3(ByVal wselect As Integer, ByVal w3page As Integer)
    Dim firstwell As Integer
      
    firstwell = 3 * w3page + 1
    
    If wselect = 2 Then
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        
        
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("D29") = "W-" & CStr(firstwell)
        
        ActiveSheet.Range("G12") = "W-" & CStr(firstwell + 1)
        ActiveSheet.Range("H29") = "W-" & CStr(firstwell + 1)
    Else
    
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
        Selection.Delete
    
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    End If
End Sub

Sub DuplicateQ3Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q3")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .themeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ3(i)
    Next i
        
End Sub

Sub DuplicateRestQ3(ByVal wselect As Integer, ByVal w3page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q" & CStr(wselect))
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.name = "p" & CStr(w3page + 1)
    
    With ActiveSheet.Tab
        .themeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRestQ3(wselect, w3page)
    
End Sub



'******************************************************************************************************************************


