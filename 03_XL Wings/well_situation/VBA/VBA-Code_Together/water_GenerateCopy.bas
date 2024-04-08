
' ***************************************************************
' water_GenerationCopy
'
' ***************************************************************

Option Explicit

Private Function LastRowByKey(cell As String) As Long
    LastRowByKey = Range(cell).End(xlDown).row
End Function


Private Function lastRowByRowsCount(cell As String) As Long
    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).row
End Function

Public Sub clearRowA()
    
End Sub

Private Function lastRowByFind() As Long
    Dim lastRow As Long
    
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastRow
End Function

Private Sub DoCopy(lastRow As Long)
    Range("F2:H" & lastRow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' 물량
    Range("L2:L" & lastRow).Select
    Selection.Copy
    
    Range("q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("k2:k" & lastRow).Select
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
    
    Num_Column = Cell_Add.column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function


' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA



Sub MainMoudleGenerateCopy()
    Dim lastRow As Long
        
    lastRow = LastRowByKey("A1")
    Call DoCopy(lastRow)
End Sub

Sub BeepExample()
    ' Make the system beep
    Beep
End Sub


'
'***************************************************************************************************************************************
'
'
'' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA

Sub ToggleOX()
    Dim activeCellColumn As String
    Dim activeCellRow As String
    Dim cp As String
    Dim fillRange As String
    Dim lastRow As Long
    Dim ret As Variant
    Dim yongdo_s As String
    
    ' Get the column and row of the active cell
    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
    
    ' Toggle "O" and "X" in column S
    If activeCellColumn = "S" Then
        If ActiveCell.Value = "O" Then
            ActiveCell.Value = "X"
        Else
            ActiveCell.Value = "O"
        End If
    End If
    
    ' Toggle between "신고공" and "허가공" in column B
    If activeCellColumn = "B" Then
        If ActiveCell.Value = "신고공" Then
            ActiveCell.Value = "허가공"
            ToggleFontSettings ActiveCell, True
        Else
            ActiveCell.Value = "신고공"
            ToggleFontSettings ActiveCell, False
        End If
    End If
    
    ' AutoFill for columns C and D
    If activeCellColumn = "C" Or activeCellColumn = "D" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastRow = LastRowByKey(ActiveCell.address)
        fillRange = activeCellColumn & Range(cp).row & ":" & activeCellColumn & lastRow
        Range(cp).AutoFill Destination:=Range(fillRange)
    End If
    
    
    
    If activeCellColumn = "H" Then
        If ActiveCell.Value > 1 Then
            ActiveCell.Value = ActiveCell.Value / 10
        Else
            ActiveCell.Value = ActiveCell.Value * 10
        End If
    End If
    
    ' Populate columns F to J based on "get_wellinfo_function()" result
    If activeCellColumn >= "F" And activeCellColumn <= "G" Then
        If IsEmpty(ActiveCell.Value) Then
            ret = get_wellinfo_function(2)
                        
            Select Case ret(0)
                Case "농업용"
                Case "농어업용"
                        yongdo_s = "aa"
    
                Case "생활용"
                        yongdo_s = "ss"
    
                Case "공업용"
                        yongdo_s = "ii"
    
                Case Else
                        yongdo_s = "ss"
            End Select
            
            If yongdo_s <> ActiveSheet.name Then
                Beep
                Debug.Print ret(0)
            Else
                PopulateWellInfo ret, activeCellRow, activeCellColumn
            End If
        End If
    End If
    
    ' Show different user forms based on the active sheet name and column K
    If activeCellColumn = "K" Then
        Select Case ActiveSheet.name
            Case "ss": UserForm_SS.Show
            Case "aa": UserForm_AA.Show
            Case "ii": UserForm_II.Show
        End Select
    End If
    
End Sub

Sub ToggleFontSettings(cell As Range, isBold As Boolean)
    With cell.Font
        If isBold Then
            .Color = -16776961
            .TintAndShade = 0
            .Bold = True
        Else
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .Bold = False
        End If
    End With
End Sub


Sub PopulateWellInfo(ret As Variant, row As String, column As String)
    Cells(row, "K").Value = ret(1) ' sebu
    Cells(row, "F").Value = ret(2) ' simdo
    Cells(row, "G").Value = ret(3) ' well_diameter
    Cells(row, "H").Value = ret(4) ' well_hp
    Cells(row, "I").Value = ret(5) ' well_q
    Cells(row, "J").Value = ret(6) ' well_tochul
End Sub




'
'***************************************************************************************************************************************
'
'


Sub SubModuleInitialClear()
    Dim lastRow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastRow = LastRowByKey("A1")
  
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    Range("e2:j" & lastRow).Select
    Selection.ClearContents
    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    
    If lastRow >= 23 Then
        Rows("23:" & lastRow).Select
        Selection.Delete Shift:=xlUp
    End If
    
    
    If (ActiveSheet.name = "ii") Then
        Range("l2").Value = 0
    End If
    
    Range("m2").Select
End Sub


Sub Finallize()
    Dim lastRow As Long
    Dim delStartRow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastRow = LastRowByKey("A1")
    delStartRow = LastRowByKey("E1") + 1
    
    
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    If delStartRow = 1048577 Or lastRow = 2 Then
        Exit Sub
    Else
        Rows(delStartRow & ":" & lastRow).Select
        Selection.Delete Shift:=xlUp
        Range("A2").Select
    End If
      
End Sub

Sub SubModuleCleanCopySection()
    Dim lastRow As Long
        
    lastRow = LastRowByKey("A1")
    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    Range("P14").Select
End Sub


' 2023/4/19 - copy modify

Sub insertRow()
    Dim lastRow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    Dim AddingRowCount As Long
    
    'lastRow = lastRowByKey("A1")

    AddingRowCount = 10

    lastRow = lastRowByRowsCount("A")
    
    Rows(CStr(lastRow + 1) & ":" & CStr(lastRow + AddingRowCount)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = LastRowByKey("A1")
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





