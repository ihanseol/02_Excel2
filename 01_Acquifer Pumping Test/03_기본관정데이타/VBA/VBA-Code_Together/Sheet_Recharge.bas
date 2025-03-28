Private Sub CommandButton1_Click()
    Call getMotorPower
    Call FindMaxMin
End Sub

Private Sub CommandButton2_Click()
    Rows("36:90").Select
    Selection.Delete Shift:=xlUp
    Range("g34").Select
End Sub

Private Sub FindMaxMin()

    Dim nWell As Integer
    Dim maxVal, minVal As Double
    Dim qMax, qMin As Double
    
    
    nWell = GetNumberOfWell()
    
    If nWell <= 1 Then
        Exit Sub
    End If
    
    maxVal = Application.WorksheetFunction.max(Range("B46:" & ColumnNumberToLetter(nWell + 1) & "46"))
    minVal = Application.WorksheetFunction.min(Range("B46:" & ColumnNumberToLetter(nWell + 1) & "46"))
    
    qMax = Application.WorksheetFunction.max(Range("B41:" & ColumnNumberToLetter(nWell + 1) & "41"))
    qMin = Application.WorksheetFunction.min(Range("B41:" & ColumnNumberToLetter(nWell + 1) & "41"))
    
    
    Range("k52") = minVal
    Range("k53") = maxVal
    
    Range("l52") = qMin
    Range("l53") = qMax

End Sub

Private Sub ShowLocation_Click()
      Sheets("location").Visible = True
      Sheets("location").Activate
End Sub


Sub hfill_red(ByVal i As Integer)
    Range("C" & i & ":P" & i).Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Sub hfill_clear()
    Range("C5:P17").Select
    With Selection.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


Private Sub CommandButton3_Click()
    Dim i As Integer
    Dim max, min As Single
    
    max = Range("o15").value
    min = Range("o16").value
    
    Range("B5:P14").Select
    Selection.Font.Bold = False
     
    Range("a1").Activate
    Call hfill_clear
    
    For i = 5 To 14
        If Cells(i, "O").value = max Or Cells(i, "O").value = min Then
            Call hfill_red(i)
            Union(Cells(i, "B"), Cells(i, "O")).Select
            Selection.Font.Bold = True
        End If
    Next i
End Sub


