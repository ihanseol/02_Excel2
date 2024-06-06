Private Sub CommandButton1_Click()
    Call importRainfall
End Sub

Private Sub CommandButton2_Click()
    Range("b5:n34").ClearContents
End Sub


Private Sub importRainfall()
    Dim myArray As Variant
    Dim rng As Range

    Select Case UCase(Range("T6").Value)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
    End Select

    Dim indexString As String
    indexString = "data_" & UCase(Range("T6").Value)

    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.Value = myArray

    Range("B2").Value = Range("T5").Value & "±‚ªÛ√ª"
End Sub







