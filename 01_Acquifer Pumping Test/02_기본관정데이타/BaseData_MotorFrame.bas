Attribute VB_Name = "BaseData_MotorFrame"
Public Sub draw_motor_frame(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    Debug.Print lastRow()
    mychar = Chr(65 + nof_sheets)
    
    Range("A" & CStr(po) & ":" & mychar & CStr(po + 12)).Select
    
    Call draw_border
    Range("A" & CStr(po) & ":" & mychar & CStr(po + 1)).Select
    Call draw_border
    Range("A" & CStr(po + 11) & ":" & mychar & CStr(po + 12)).Select
    Call draw_border
    Range("A" & CStr(po) & ":" & "A" & CStr(po + 12)).Select
    Call draw_border
    
    Range("A" & CStr(po + 2) & ":" & "A" & CStr(po + 10)).Select
    Call draw_border
    
    Range("A" & CStr(po) & ":B" & CStr(po)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Range("a" & CStr(po)).value = "�������»����ϴ°�"
    Range("a" & CStr(po + 2)).value = "�����ɵ�"
    Range("a" & CStr(po + 3)).value = "Q(����)-�����"
    Range("a" & CStr(po + 4)).value = "Depth(���ͼ�ġ�ɵ�)"
    Range("a" & CStr(po + 5)).value = "Height(������)"
    Range("a" & CStr(po + 6)).value = "Sum (�հ�)"
    Range("a" & CStr(po + 7)).value = "E (ȿ��)"
    Range("a" & CStr(po + 9)).value = "����"
    Range("a" & CStr(po + 11)).value = "�㰡������ ����"
    Range("a" & CStr(po + 12)).value = "�̷л� ����ɷ�"
    
    Call decorationPumpHP(nof_sheets, po)
    Call decorationInerLine(nof_sheets, po)
    Call alignTitle(nof_sheets, po)
End Sub

Private Sub alignTitle(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    mychar = Chr(65 + nof_sheets)
    
    'Range("A57:B57").Select
    Range("A" & CStr(po) & ":" & "B" & CStr(po)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Selection.Font.Italic = False
    With Selection.Font
        .name = "���� ���"
        .Size = 12
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("A59:A69").Select
    Range("A" & CStr(po + 2) & ":" & "A" & CStr(po + 12)).Select
    
    With Selection.Font
        .name = "���� ���"
        .Size = 11
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    Range("O65").Select
End Sub

Private Sub decorationPumpHP(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    mychar = Chr(65 + nof_sheets)
    'Range("B58:N69").Select
    Range("B" & CStr(po + 1) & ":" & mychar & CStr(po + 12)).Select
    
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .name = "���� ���"
        .Size = 12
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("B59:N59").Select
    Range("B" & CStr(po + 2) & ":" & mychar & CStr(po + 2)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    'Range("B60:N60").Select
    Range("B" & CStr(po + 3) & ":" & mychar & CStr(po + 3)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    
    'Range("B63:N63").Select
    Range("B" & CStr(po + 6) & ":" & mychar & CStr(po + 6)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    ActiveWindow.SmallScroll Down:=3
    
    'Range("B64:N64").Select
    Range("B" & CStr(po + 7) & ":" & mychar & CStr(po + 7)).Select
    Selection.NumberFormatLocal = "0.00"
    
    'Range("B67:N67").Select
    Range("B" & CStr(po + 10) & ":" & mychar & CStr(po + 10)).Select
    Selection.Font.Bold = True
    With Selection.Font
        .name = "���� ���"
        .Size = 14
        .Italic = True
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("B68:N69").Select
    Range("B" & CStr(po + 11) & ":" & mychar & CStr(po + 12)).Select
    
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub

Private Sub decorationInerLine(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    mychar = Chr(65 + nof_sheets)
    'Range("A60:N61").Select
    Range("A" & CStr(po + 3) & ":" & mychar & CStr(po + 4)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'Range("B67:N67").Select
    Range("B" & CStr(po + 10) & ":" & mychar & CStr(po + 10)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'Range("B59:N67").Select
    Range("B" & CStr(po + 2) & ":" & mychar & CStr(po + 10)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("A59:A67").Select
    Range("A" & CStr(po + 2) & ":" & "A" & CStr(po + 10)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("B68:N69").Select
    Range("B" & CStr(po + 11) & ":" & mychar & CStr(po + 12)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("A68:A69").Select
    Range("A" & CStr(po + 11) & ":" & "A" & CStr(po + 12)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
End Sub

Private Sub draw_border()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
