'
' Refactor By User Defined Type
'


' Define User-Defined Type for Well Data
Private Type WellData
    Natural As Double
    Stable As Double
    Recover As Double
    
    DeltaH As Double
    Sw As Double
    Radius As Double
    
    Rw As Double
    WellDepth As Double
    Casing As Double
    
    Q As Double
    DeltaS As Double
    Hp As Double
    DaeSoo As Double
    
    T1 As Double
    T2 As Double
    TA As Double
    S1 As Double
    S2 As Double
    
    K As Double
    Time As Double
    
    Shultze As Double
    Webber As Double
    Jacob As Double
    Skin As Double
    
    Er As Double
    ER1 As Double
    ER2 As Double
    ER3 As Double
    
    Qh As Double
    Qg As Double
    Sd1 As Double
    Sd2 As Double
    Q1 As Double
    
    C As Double
    B As Double
    Ratio As Double
    
    T0 As Double
    S0 As Double
    ERMode As String
    Address As String
    Company As String
    
    S3 As Double
    Title As String
End Type

' Main procedure using UDT
Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim nofwell As Integer
    Dim i As Integer
    Dim rngString As String
    
    ' Get total number of wells
    nofwell = GetNumberOfWell()
    
    ' Determine range to clear based on import type
    If Not isSingleWellImport And singleWell = 999 Then
        rngString = "A5:AR37"
    Else
        rngString = "A" & (nofwell + 4) & ":AR" & (nofwell + 4)
    End If
    
    Call EraseCellData(rngString)
    
    ' Process each well
    For i = 1 To nofwell
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            Dim well As WellData
            well = ImportDataForWell(i)
            SetWellDataToSheet i, well
        End If
    Next i
End Sub

' Function to import well data into UDT
Private Function ImportDataForWell(ByVal wellIndex As Integer) As WellData
    Dim fName As String
    Dim wb As Workbook
    Dim well As WellData
    
    ' Construct filename and check if workbook is open
    fName = "A" & CStr(wellIndex) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fName) Then
        MsgBox "Please open the yangsoo data! " & fName
        Exit Function
    End If
    
    Set wb = Workbooks(fName)
    
    ' Get worksheet references
    With wb
        Dim wsInput As Worksheet: Set wsInput = .Worksheets("Input")
        Dim wsSkinFactor As Worksheet: Set wsSkinFactor = .Worksheets("SkinFactor")
        Dim wsSafeYield As Worksheet: Set wsSafeYield = .Worksheets("SafeYield")
        
        ' Populate UDT with data
        With well
            .Natural = wsInput.Range("m48").value ' 자연수위
            .Stable = wsInput.Range("m49").value ' 안정수위
            .Recover = wsSkinFactor.Range("c10").value ' 회복수위
            .Sw = wsSkinFactor.Range("c11").value ' 회복수위
            
            .DeltaH = wsSkinFactor.Range("b16").value  ' 수위강하량
            .Radius = wsInput.Range("m44").value '200 , 우물직경
            .Rw = wsSkinFactor.Range("e4").value '0.1
            
            .WellDepth = wsInput.Range("m45").value ' 관정심도
            .Casing = wsInput.Range("i52").value ' 케이싱심도
            .Q = wsInput.Range("m51").value '양수량, 채수계획량
            .C = wsInput.Range("a31").value ' 우물손실계수
            .B = wsInput.Range("b31").value ' 대수층손실계수
            
            .DeltaS = wsSkinFactor.Range("b4").value '최초 1분간 수위강하량
            .Hp = wsInput.Range("i48").value  ' 모터마력
            .DaeSoo = wsSkinFactor.Range("c16").value ' 대수층두께
            
            .T1 = wsSkinFactor.Range("d5").value
            .T2 = wsSkinFactor.Range("h13").value
            .TA = wsSkinFactor.Range("d16").value
            .S1 = wsSkinFactor.Range("e10").value
            .S2 = wsSkinFactor.Range("i16").value
            
            .K = wsSkinFactor.Range("e16").value
            .Time = wsSkinFactor.Range("h16").value
            
            .Shultze = wsSkinFactor.Range("c13").value
            .Webber = wsSkinFactor.Range("c18").value
            .Jacob = wsSkinFactor.Range("c23").value
            
            .Skin = wsSkinFactor.Range("g6").value
            .Er = wsSkinFactor.Range("c8").value
            
            .ER1 = wsSkinFactor.Range("k8").value
            .ER2 = wsSkinFactor.Range("k9").value
            .ER3 = wsSkinFactor.Range("k10").value
            
            .Qh = wsInput.Range("d6").value
            .Qg = wsSafeYield.Range("b8").value
            
            .Sd1 = wsSafeYield.Range("b3").value
            .Sd2 = wsSafeYield.Range("b4").value
            .Q1 = wsSafeYield.Range("b2").value
            .Ratio = wsSafeYield.Range("b11").value
            
            .T0 = wsSkinFactor.Range("d4").value
            .S0 = wsSkinFactor.Range("f4").value
            
            .ERMode = wsSkinFactor.Range("h10").value
            .Address = wsInput.Range("i46").value
            .Company = wsInput.Range("i47").value
            
            .S3 = wsSkinFactor.Range("i13").value
            .Title = wsInput.Range("i44").value
        End With
    End With
    
    ImportDataForWell = well
End Function

' Procedure to set well data to worksheet
Private Sub SetWellDataToSheet(ByVal wellIndex As Integer, well As WellData)
    Dim row As Long: row = 4 + wellIndex
    Cells(row, 1).value = "W-" & wellIndex
    
    With well
        SetCellValue row, 2, .Natural, "0.00"
        SetCellValue row, 3, .Stable, "0.00"
        SetCellValue row, 4, .Recover, "0.00"
        SetCellValue row, 5, .Sw, "0.00"
        
        SetCellValue row, 6, .DeltaH, "0.00"
        SetCellValue row, 7, .Radius, "0."
        
        SetCellValue row, 8, .Rw, "0.000"
        SetCellValue row, 9, .WellDepth, "0"
        
        SetCellValue row, 10, .Casing, "0"
        SetCellValue row, 11, .Q, "0"
        SetCellValue row, 12, .DeltaS, "0.00"
        SetCellValue row, 13, .Hp, "0.0"
        SetCellValue row, 14, .DaeSoo, "0.00"
        
        SetCellValue row, 15, .T1, "0.0000"
        SetCellValue row, 16, .T2, "0.0000"
        SetCellValue row, 17, .TA, "0.0000"
        
        SetCellValue row, 18, .S1, "0.0000000"
        SetCellValue row, 19, .S2, "0.0000000"
        
        SetCellValue row, 20, .K, "0.0000"
        SetCellValue row, 21, .Time, "0.0000"
        SetCellValue row, 22, .Shultze, "0.00"
        SetCellValue row, 23, .Webber, "0.00"
        SetCellValue row, 24, .Jacob, "0.00"
        
        SetCellValue row, 25, .Skin, "0.0000"
        SetCellValue row, 26, .Er, "0.0000"
        SetCellValue row, 27, .Qh, "0"
        SetCellValue row, 28, .Qg, "0.00"
        
        SetCellValue row, 29, .Q1, "0.00"
        SetCellValue row, 30, .Sd1, "0.00"
        SetCellValue row, 31, .Sd2, "0.00"
        
        SetCellValue row, 32, .C, "0.00"
        SetCellValue row, 33, .B, "0.00"
        SetCellValue row, 34, .Ratio, "0.0%"
        
        SetCellValue row, 35, .T0, "0.0000"
        SetCellValue row, 36, .S0, "0.0000"
        
        SetCellValue row, 37, .ERMode, ""
        SetCellValue row, 38, .ER1, "0.0000"
        SetCellValue row, 39, .ER2, "0.0000"
        SetCellValue row, 40, .ER3, "0.0000"
        
        SetCellValue row, 41, .Address, ""
        SetCellValue row, 42, .Company, ""
        SetCellValue row, 43, .S3, "0.00"
        SetCellValue row, 44, .Title, ""
    End With
End Sub

' Helper procedure to set cell value and format
Private Sub SetCellValue(ByVal row As Long, ByVal col As Integer, ByVal value As Variant, ByVal numberFormat As String)
    With Cells(row, col)
        .value = value
        If Len(numberFormat) > 0 Then .numberFormat = numberFormat
    End With
End Sub

' Keep existing helper functions
Public Sub MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True
      DoEvents
      Debug.Print String(30, vbCrLf)
   End If
   Debug.Print sPrintStr
End Sub

Function DetermineEffectiveRadius(ERMode As String) As Integer
    Dim Er As String, R As String
    Er = ERMode
    R = Mid(Er, 5, 1)
    
    If R = "F" Then
        DetermineEffectiveRadius = 0 ' Assuming erRE0 = 0, adjust if different
    Else
        DetermineEffectiveRadius = val(R)
    End If
End Function

Function CheckFileExistence(filePath As String) As Boolean
    CheckFileExistence = (Dir(filePath) <> "")
End Function



Sub WriteFormula()
 
    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double
    
    ' Save array to a file
    Dim filePath As String
    Dim FileNum As Integer
    
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
    
    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    DeleteFileIfExists filePath
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
        
    Call MyDebug("Formula SkinFactor ... ", True)
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    ' 스킨계수
    Call FormulaSkinFactorAndER("SKIN", FileNum)
    
    ' 유효우물반경
    Call FormulaSkinFactorAndER("ER", FileNum)
    
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    Call FormulaChwiSoo(FileNum)
    ' 3-7, 적정취수량
    
    Call FormulaRadiusOfInfluence(FileNum)
    ' 양수영향반경
        
    Close FileNum

End Sub



Sub FormulaSkinFactorAndER(ByVal Mode As String, ByVal FileNum As Integer)
    Dim formula1, formula2 As String
    Dim nofwell As Integer
    Dim i As Integer
    Dim T, Q, Radius, Skin, Er As Double
    Dim T0, S0 As Double
    Dim ERMode As String
    Dim ER1, ER2, ER3, B, S1 As Double
    
        
    Call MyDebug("Formula SkinFactor ... ", True)
    
    nofwell = GetNumberOfWell()
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        T = format(Cells(4 + i, "o").value, "0.0000")
        Q = Cells(4 + i, "k").value
        
        T0 = format(Cells(4 + i, "AI").value, "0.0000")
        S0 = format(Cells(4 + i, "AJ").value, "0.0000")
        S1 = Cells(4 + i, "R").value
                
        delta_s = format(Cells(4 + i, "l").value, "0.00")
        Radius = format(Cells(4 + i, "h").value, "0.000")
        Skin = format(Cells(4 + i, "y").value, "0.0000")
        Er = format(Cells(4 + i, "z").value, "0.0000")
        
        
        B = format(Cells(4 + i, "AG").value, "0.0000")
        ER1 = format(Cells(4 + i, "AL").value, "0.0000")
        ER2 = format(Cells(4 + i, "AM").value, "0.0000")
        ER3 = format(Cells(4 + i, "AN").value, "0.0000")
        
        
        Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
        ' 경험식 1번
        Case erRE1
            
            formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
            formula2 = "erRE1, 경험식 1번"
            
        ' 경험식 2번
        Case erRE2
            formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
            formula2 = "erRE2, 경험식 2번"
        ' 경험식 3번
        Case erRE3
            formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~" & Radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
            formula2 = "erRE3, 경험식 3번"
            
        Case Else
            ' 스킨계수
            formula1 = "W-" & i & "호공~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {" & S0 & " TIMES  (" & Radius & " TIMES  " & Radius & ")} =`" & Skin
            ' 유효우물반경
            formula2 = "W-" & i & "호공~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & Radius & " TIMES e ^{-(" & Skin & ")} =" & Er & "m"
        End Select
        
        
        If Mode = "SKIN" Then
            Debug.Print formula1
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            
            Print #FileNum, formula1
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Else
            Debug.Print formula2
            Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                
            Print #FileNum, formula2
            Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        End If
    Next i

End Sub


Sub DeleteFileIfExists(filePath As String)
    If Len(Dir(filePath)) > 0 Then ' Check if file exists
        On Error Resume Next
        Kill filePath ' Attempt to delete the file
        
        On Error GoTo 0
        If Len(Dir(filePath)) > 0 Then
            MsgBox "Unable to delete file '" & filePath & "'. The file may be in use.", vbExclamation
        Else
            ' MsgBox "File '" & filePath & "' has been deleted.", vbInformation
            
        End If
    Else
        ' MsgBox "File '" & filePath & "' does not exist.", vbExclamation
    End If
End Sub



Sub FormulaChwiSoo(FileNum As Integer)
' 3-7, 적정취수량

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim Q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        Q1 = Cells(4 + i, "ac").value
 
        S1 = format(Cells(4 + i, "ad").value, "0.00")
        S2 = format(Cells(4 + i, "ae").value, "0.00")
        res = format(Cells(4 + i, "ab").value, "0.00")
        
        'W-1호공~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =165.52㎥/일
        'W-1호공~~Q _{& 2} =100 TIMES  ( {8.71} over {4.09} ) ^{2/3} =`165.52㎥/일
        
        formula = "W-" & i & "호공~~Q_{& 2} = " & Q1 & " TIMES ({" & S2 & "} over {" & S1 & "}) ^{2/3} = `" & res & " ㎥/일"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
      
End Sub


Sub FormulaRadiusOfInfluence(FileNum As Integer)
' 양수영향반경

    Call FormulaSUB_ROI("SCHULTZE", FileNum)
    Call FormulaSUB_ROI("WEBBER", FileNum)
    Call FormulaSUB_ROI("JCOB", FileNum)
    
End Sub




Sub FormulaSUB_ROI(Mode As String, FileNum As Integer)
  Dim formula1, formula2, formula3 As String
    ' 슐츠, 웨버, 제이콥
    
    Dim nofwell As String
    Dim i As Integer
    Dim Shultze, Webber, Jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        schultze = CStr(format(Cells(4 + i, "v").value, "0.0"))
        Webber = CStr(format(Cells(4 + i, "w").value, "0.0"))
        Jacob = CStr(format(Cells(4 + i, "x").value, "0.0"))
        
        T = CStr(format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(format(Cells(4 + i, "t").value, "0.0000"))
    
        delta_h = CStr(format(Cells(4 + i, "f").value, "0.00"))
        time_ = CStr(format(Cells(4 + i, "u").value, "0.0000"))
        
        
        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        formula1 = "W-" & i & "호공~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & schultze & "m"
        formula2 = "W-" & i & "호공~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & Webber & "`m"
        formula3 = "W-" & i & "호공~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & Jacob & "m"
        
        
        Select Case Mode
            Case "SCHULTZE"
                Debug.Print formula1
                Print #FileNum, formula1
            
            Case "WEBBER"
                Debug.Print formula2
                Print #FileNum, formula2
                
            Case "JCOB"
                Debug.Print formula3
                Print #FileNum, formula3
        End Select
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
End Sub





