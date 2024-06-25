Option Explicit



Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("AggStep").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
'Collect Data
    
    Call WriteStepTestData(999, False)
End Sub



Private Sub CommandButton3_Click()
'Single Well Import

'single well import

Dim singleWell  As Integer
Dim WB_NAME As String


WB_NAME = GetOtherFileName
'MsgBox WB_NAME

'If Workbook Is Nothing Then
'    GetOtherFileName = "Empty"
'Else
'    GetOtherFileName = Workbook.name
'End If
    
If WB_NAME = "Empty" Then
    MsgBox "WorkBook is Empty"
    Exit Sub
Else
    singleWell = CInt(ExtractNumberFromString(WB_NAME))
'   MsgBox (SingleWell)
End If

Call WriteStepTestData(singleWell, True)

End Sub


