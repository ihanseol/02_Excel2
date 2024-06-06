Option Explicit

Private Sub CommandButton1_Click()
    Sheets("aggWhpa").Visible = False
    Sheets("Well").Select
End Sub


Function getDirectionFromWell(i) As Integer

    If Sheets(CStr(i)).Range("k12").Font.Bold Then
        getDirectionFromWell = Sheets(CStr(i)).Range("k12").Value
    Else
        getDirectionFromWell = Sheets(CStr(i)).Range("l12").Value
    End If

End Function


Private Sub CommandButton2_Click()
' Collect Data

    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Q, daeSoo, T1, S1, direction, gradient As Double
    
    nofwell = sheets_count()
    If ActiveSheet.name <> "aggWhpa" Then Sheets("aggWhpa").Select
    Call EraseCellData("C4:O34")
    
    TurnOffStuff
    
    For i = 1 To nofwell
        Q = Sheets(CStr(i)).Range("c16").Value
        daeSoo = Sheets(CStr(i)).Range("c14").Value
        
        T1 = Sheets(CStr(i)).Range("e7").Value
        S1 = Sheets(CStr(i)).Range("g7").Value
        
        direction = getDirectionFromWell(i)
        gradient = Sheets(CStr(i)).Range("k18").Value
        
        Call WriteWellData_Single(Q, daeSoo, T1, S1, direction, gradient, i)
    Next i
    
    Sheets("aggWhpa").Select
    
    Call MakeAverageAndMergeCells(nofwell)
    Call DrawOutline
    TurnOnStuff
    
End Sub

Private Sub WriteWellData_Single(Q As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, direction As Variant, gradient As Variant, ByVal i As Integer)
    
    Call UnmergeAllCells
        
    Cells(3 + i, "c").Value = "W-" & CStr(i)
    Cells(3 + i, "e").Value = Q
    Cells(3 + i, "f").Value = T1
    Cells(3 + i, "i").Value = daeSoo
    Cells(3 + i, "k").Value = direction
    Cells(3 + i, "m").Value = Format(gradient, "###0.0000")
    Cells(4, "d").Value = "5��"
    
End Sub


Sub MakeAverageAndMergeCells(ByVal nofwell As Integer)
    Dim t_sum, daesoo_sum, gradient_sum, direction_sum As Double
    Dim i As Integer

    For i = 1 To nofwell
        t_sum = t_sum + Range("F" & (i + 3)).Value
        daesoo_sum = daesoo_sum + Range("I" & (i + 3)).Value
        direction_sum = direction_sum + Range("K" & (i + 3)).Value
        gradient_sum = gradient_sum + Range("M" & (i + 3)).Value
    Next i
    
    
    Cells(4, "g").Value = Round(t_sum / nofwell, 4)
    Cells(4, "g").NumberFormat = "0.0000"
    
    Cells(4, "j").Value = Round(daesoo_sum / nofwell, 1)
    Cells(4, "j").NumberFormat = "0.0"
        
    Cells(4, "l").Value = Round(direction_sum / nofwell, 1)
    Cells(4, "l").NumberFormat = "0.0"
        
    Cells(4, "n").Value = Round(gradient_sum / nofwell, 4)
    Cells(4, "n").NumberFormat = "0.0000"
       
    Cells(4, "o").Value = "���������"
    Cells(4, "h").Value = 0.03
    
    Call merge_cells("d", nofwell)
    Call merge_cells("g", nofwell)
    Call merge_cells("j", nofwell)
    Call merge_cells("l", nofwell)
    Call merge_cells("n", nofwell)
    Call merge_cells("o", nofwell)
    Call merge_cells("h", nofwell)

End Sub


Sub merge_cells(cel As String, ByVal nofwell As Integer)

    Range(cel & CStr(4) & ":" & cel & CStr(nofwell + 3)).Select
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
    
End Sub

Sub FindMergedCellsRange()
    Dim mergedRange As Range
    Set mergedRange = ActiveSheet.Range("A1").MergeArea
    MsgBox "The range of merged cells is " & mergedRange.Address
End Sub



Sub UnmergeAllCells()
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = ActiveSheet
    
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            cell.MergeCells = False
        End If
    Next cell
    
    Call SetBorderLine_Default
End Sub


Sub SetBorderLine_Default()

    Range("C4:O17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .Value = ""
    End With
End Sub


Sub DrawOutline()

    Application.ScreenUpdating = False
    
    Range("C3:O34").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B31").Select
    
    Application.ScreenUpdating = True
End Sub






