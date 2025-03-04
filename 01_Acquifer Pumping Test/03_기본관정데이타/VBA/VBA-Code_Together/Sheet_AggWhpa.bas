Option Explicit

Private Sub CommandButton1_Click()
    Sheets("aggWhpa").Visible = False
    Sheets("Well").Select
End Sub



Private Sub CommandButton2_Click()
' Collect Data

    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Q, DaeSoo, T1, S1, direction, gradient As Double
    
    nofwell = sheets_count()
    If ActiveSheet.name <> "aggWhpa" Then Sheets("aggWhpa").Select
    Call EraseCellData("C4:O34")
    
    TurnOffStuff
    
    For i = 1 To nofwell
        Q = Sheets(CStr(i)).Range("c16").value
        DaeSoo = Sheets(CStr(i)).Range("c14").value
        
        T1 = Sheets(CStr(i)).Range("e7").value
        S1 = Sheets(CStr(i)).Range("g7").value
        
        direction = getDirectionFromWell(i)
        gradient = Sheets(CStr(i)).Range("k18").value
        
        Call modAggWhpa.WriteWellData_Single(Q, DaeSoo, T1, S1, direction, gradient, i)
    Next i
    
    Sheets("aggWhpa").Select
    
    Call MakeAverageAndMergeCells(nofwell)
    Call DrawOutline
    TurnOnStuff
    
End Sub





