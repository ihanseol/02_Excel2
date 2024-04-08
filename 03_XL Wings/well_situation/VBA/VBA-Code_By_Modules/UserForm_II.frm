VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_II 
   Caption         =   "II, Contents Selection"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405.001
   OleObjectBlob   =   "UserForm_II.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_II"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ***************************************************************
' UserForm_II
'
' ***************************************************************

' Optionbutton1 - 자유입지업체
' Optionbutton2 - 기타
' Optionbutton3 - 지방공단
' Optionbutton4 - 농공단지


Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("자유입지업체", "기타", "지방공단", "농공단지")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 3
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub

Private Sub CommandButton2_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
End Sub

'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 Then
'        Unload Me
'    End If
'End Sub

