Sub ��ũ��1()
'
' ��ũ��1 ��ũ��
'

'
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
End Sub


Sub Step_Pumping_Test()
    Dim i           As Integer
    
    Application.ScreenUpdating = False
    
    ' ----------------------------------------------------------------
    
    Range("D3:D7").Select
    '����, Q
    
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
    ' ���ϼ���
    
    Selection.Copy
    Range("R44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    
    ' ----------------------------------------------------------------
    
    Range("B3:B7").Select
    'Sw, ���ϼ���
    
    Selection.Copy
      
    Range("S44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    ' ----------------------------------------------------------------

    Range("G3:G7").Select
    'Q/Sw , ������
    
    Selection.Copy
    Range("T44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    ' Selection.NumberFormatLocal = "0.000"
    
    Call CutDownNumber("T", 3)
     Application.CutCopyMode = False
    ' ----------------------------------------------------------------
    
    Range("F3:F7").Select
    'Sw/Q, ��������Ϸ�
    
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




