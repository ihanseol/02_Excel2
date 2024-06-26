Sub Popup_MessageBox(ByVal msg As String)
    UserForm1.TextBox1.Text = msg
    UserForm1.Show
End Sub

Sub Popup_CloseUserForm()
    Unload UserForm1
End Sub

Sub test()
    ' Application.OnTime Now + TimeValue("00:00:01"), "Popup_CloseUserForm"
    Popup_MessageBox ("Automatic Close at One Seconds ...")
End Sub


