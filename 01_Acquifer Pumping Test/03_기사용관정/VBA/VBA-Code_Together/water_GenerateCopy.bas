
' ***************************************************************
' water_GenerationCopy
'
' ***************************************************************

Option Explicit

Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


Private Function lastRowByRowsCount(cell As String) As Long
    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).row
End Function

Public Sub clearRowA()
    
End Sub

Private Function lastRowByFindAll() As Long
    Dim lastrow As Long
    
    lastrow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastrow
End Function


Private Function lastRowByFind(ByVal str As String) As Long
    Dim lastrow As Long
    
    lastrow = Cells.Find(str, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastrow
End Function


Private Sub DoCopy(lastrow As Long)
    Range("F2:H" & lastrow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' 물량
    Range("L2:L" & lastrow).Select
    Selection.Copy
    
    Range("q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("k2:k" & lastrow).Select
    Selection.Copy
    
    Range("r2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
End Sub



' return letter of range ...
Function Alpha_Column(Cell_Add As Range) As String
    Dim No_of_Rows As Integer
    Dim No_of_Cols As Integer
    Dim Num_Column As Integer
    
    No_of_Rows = Cell_Add.Rows.Count
    No_of_Cols = Cell_Add.Columns.Count
    
    If ((No_of_Rows <> 1) Or (No_of_Cols <> 1)) Then
        Alpha_Column = ""
        Exit Function
    End If
    
    Num_Column = Cell_Add.Column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function


' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA
Sub ToggleOX()
    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastrow As Long
    Dim cp, fillRange As String
    

    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Debug.Print Alpha_Column(ActiveCell)
    
    If activeCellColumn = "S" Then
        If ActiveCell.Value = "O" Then
            ActiveCell.Value = "X"
        Else
            ActiveCell.Value = "O"
        End If
    End If
    

    If activeCellColumn = "B" Then
        If ActiveCell.Value = "신고공" Then
            ActiveCell.Value = "허가공"
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        Else
            ActiveCell.Value = "신고공"
             With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If
    End If
    
    If activeCellColumn = "D" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastrow = lastRowByKey(ActiveCell.address)
        
        fillRange = "D" & Range(cp).row & ":D" & lastrow
        
        Range(cp).Select
        Selection.Copy
        Range(fillRange).Select
        
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range(cp).Select
        Application.CutCopyMode = False
    End If
    
    If activeCellColumn = "C" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastrow = lastRowByKey(ActiveCell.address)
        
        fillRange = "C" & Range(cp).row & ":C" & lastrow
        
        Range(cp).Select
        Selection.Copy
        Range(fillRange).Select
        
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Range(cp).Select
        Application.CutCopyMode = False
    End If
    
    
    If ActiveSheet.Name = "ss" And activeCellColumn = "K" Then
        UserForm_SS.Show
    End If
    
    If ActiveSheet.Name = "aa" And activeCellColumn = "K" Then
        UserForm_AA.Show
    End If
    
    If ActiveSheet.Name = "ii" And activeCellColumn = "K" Then
        UserForm_II.Show
    End If
End Sub

Sub MainMoudleGenerateCopy()
    Dim lastrow As Long
        
    lastrow = lastRowByKey("A1")
    Call DoCopy(lastrow)
End Sub


Sub SubModuleInitialClear()
    Dim lastrow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastrow = lastRowByKey("A1")
  
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    Range("e2:j" & lastrow).Select
    Selection.ClearContents
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    
    If lastrow >= 23 Then
        Rows("23:" & lastrow).Select
        Selection.Delete Shift:=xlUp
    End If
    
    
    If (ActiveSheet.Name = "ii") Then
        Range("l2").Value = 0
    End If
    
    Range("m2").Select
End Sub


Sub Finallize()
    Dim lastrow As Long
    Dim delStartRow, delEndRow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastrow = lastRowByKey("A1")
    delStartRow = lastRowByKey("D1") + 1
    
    Select Case ActiveSheet.Name
    
        Case "ss"
            delEndRow = lastRowByFind("구분") - 4
            
        Case "aa"
            delEndRow = lastRowByFind("유역내") - 4
        
        Case "ii"
            delEndRow = lastRowByFind("유역내") - 6
    
    End Select
    
    
    If Range("L2").Value = 0 Then
        delStartRow = 3
        delEndRow = lastRowByKey("L1")
    End If
    
    
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    If delStartRow = 1048577 Or lastrow = 2 Or (delEndRow - delStartRow <= 3) Then
        Exit Sub
    Else
        Rows(delStartRow & ":" & delEndRow).Select
        Selection.Delete Shift:=xlUp
        Range("A2").Select
    End If
      
End Sub

Sub SubModuleCleanCopySection()
    Dim lastrow As Long
        
    lastrow = lastRowByKey("A1")
    Range("n2:r" & lastrow).Select
    Selection.ClearContents
    Range("P14").Select
End Sub


' 2023/4/19 - copy modify

Sub insertRow()
    Dim lastrow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    Dim AddingRowCount As Long
    
    'lastRow = lastRowByKey("A1")

    AddingRowCount = 10

    lastrow = lastRowByRowsCount("A")
    
    Rows(CStr(lastrow + 1) & ":" & CStr(lastrow + AddingRowCount)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1")
    j = i + AddingRowCount
    
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "K" & i & ":M" & i
    selection_target = "K" & i & ":M" & j

    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("S" & i).Select
    Selection.AutoFill Destination:=Range("S" & i & ":S" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
End Sub





