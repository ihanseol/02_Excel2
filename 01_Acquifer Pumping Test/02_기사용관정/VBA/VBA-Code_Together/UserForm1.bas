

Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:02"), "Popup_CloseUserForm"
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
   
    Me.TextBox1.Text = "this is Sample initialize"
End Sub

