
Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:02"), "Popup_CloseUserForm"
End Sub



Private Sub UserForm_Initialize()
 Me.TextBox1.Text = "this is Sample initialize"
End Sub


