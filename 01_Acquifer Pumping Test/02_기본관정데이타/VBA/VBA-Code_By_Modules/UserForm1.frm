VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm_ShowMessage"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8325.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    Application.OnTime Now + TimeValue("00:00:02"), "Popup_CloseUserForm"
End Sub



Private Sub UserForm_Initialize()
 Me.TextBox1.Text = "this is Sample initialize"
End Sub


