Option Explicit

Private Sub CommandButton_CB1_Click()
Top:
    On Error GoTo ErrorCheck
    Call set_CB1
    Exit Sub
    
ErrorCheck:
    GoTo Top
End Sub

Private Sub CommandButton_CB2_Click()
Top:
    On Error GoTo ErrorCheck
    Call set_CB2
    Exit Sub
    
ErrorCheck:
    GoTo Top
End Sub

' Chart Button
Private Sub CommandButton_Chart_Click()
    Dim gong        As Integer
    Dim KeyCell     As Range
    
    Call adjustChartGraph
    
    Set KeyCell = Range("J48")
    
    gong = Val(CleanString(KeyCell.Value))
    Call mod_Chart.SetChartTitleText(gong)
    
    Call mod_INPUT.Step_Pumping_Test
    Call mod_INPUT.Vertical_Copy
End Sub

Private Sub CommandButton_ClearReport_Click()
    
    DeleteStepSheet
       
End Sub

Sub DeleteStepSheet()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    On Error Resume Next
    Set ws1 = Sheets("Step")
    Set ws2 = Sheets("out")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        Application.DisplayAlerts = False
        Sheets("Step").Delete
        Application.DisplayAlerts = True
    Else
        Debug.Print "Sheet 'Step' does not exist."
    End If
    
    If Not ws2 Is Nothing Then
        Application.DisplayAlerts = False
        Sheets("out").Delete
        Application.DisplayAlerts = True
    Else
        Debug.Print "Sheet 'out' does not exist."
    End If
End Sub



Private Sub CommandButton_STEP_Click()
    Call Make_Step_Document
End Sub

Private Sub CommandButton_2880_Click()
    'Call make_long_document
    Call Make2880_Document
End Sub

Private Sub CommandButton_1440_Click()
    Call Make2880_Document
    Call make1440sheet
End Sub

Private Sub CommandButton8_Click()
    Call set_CB_ALL
End Sub

Private Sub Worksheet_Activate()
    '  Dim gong     As Integer
    '  Dim KeyCell  As Range
    '
    '  Set KeyCell = Range("J48")
    '
    '  gong = Val(CleanString(KeyCell.Value))
    '  Call SetChartTitleText(gong)
End Sub
