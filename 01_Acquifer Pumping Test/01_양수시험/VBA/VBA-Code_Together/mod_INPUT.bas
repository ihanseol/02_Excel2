Sub 매크로1()
'
' 매크로1 매크로
'

'
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
End Sub


Sub CommandButton_CB1_ClickRun()
Top:
    On Error GoTo ErrorCheck
    Call set_CB1
    Exit Sub
    
ErrorCheck:
    GoTo Top
End Sub

Sub CommandButton_CB2_ClickRun()
Top:
    On Error GoTo ErrorCheck
    Call set_CB2
    Exit Sub
    
ErrorCheck:
    GoTo Top
End Sub

' Chart Button
Sub CommandButton_Chart_ClickRun()
    Dim gong        As Integer
    Dim KeyCell     As Range
    
    Call adjustChartGraph
    
    Set KeyCell = Range("J48")
    
    gong = Val(CleanString(KeyCell.Value))
    Call mod_Chart.SetChartTitleText(gong)
    
    Call mod_INPUT.Step_Pumping_Test
    Call mod_INPUT.Vertical_Copy
End Sub



Sub Step_Pumping_Test()
    Dim i           As Integer
    
    Application.ScreenUpdating = False
    
    ' ----------------------------------------------------------------
    
    Range("D3:D7").Select
    '물량, Q
    
    Selection.Copy
    
    Range("Q44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0"
    
    Range("Q44:Q48").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    
    Call CutDownNumber("Q", 0)
    
    ' ----------------------------------------------------------------
    
    Range("A3:A7").Select
    ' 지하수위
    
    Selection.Copy
    Range("R44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    
    ' ----------------------------------------------------------------
    
    Range("B3:B7").Select
    'Sw, 강하수위
    
    Selection.Copy
      
    Range("S44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    ' ----------------------------------------------------------------

    Range("G3:G7").Select
    'Q/Sw , 비양수량
    
    Selection.Copy
    Range("T44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    ' Selection.NumberFormatLocal = "0.000"
    
    Call CutDownNumber("T", 3)
     Application.CutCopyMode = False
    ' ----------------------------------------------------------------
    
    Range("F3:F7").Select
    'Sw/Q, 비수위강하량
    
    Selection.Copy
    Range("U44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Call CutDownNumber("T", 5)
    
    'Range("T44:T48").Select
    'Selection.NumberFormatLocal = "0.000"
    
    ' ----------------------------------------------------------------
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

Sub Vertical_Copy()
    Dim strValue(1 To 5) As String
    Dim result As String
    Dim i As Integer
    
    EraseCellData ("q64:x64")
    
    strValue(1) = "Q44:Q48"
    strValue(2) = "R44:R48"
    strValue(3) = "S44:S48"
    strValue(4) = "T44:T48"
    strValue(5) = "U44:U48"

    For i = 1 To 5
        result = ConcatenateCells(strValue(i))
        Cells(64, Chr(81 + i - 1)).Value = result
    Next i
    
    Range("q63").Select
    Call WriteStringEtc
End Sub






