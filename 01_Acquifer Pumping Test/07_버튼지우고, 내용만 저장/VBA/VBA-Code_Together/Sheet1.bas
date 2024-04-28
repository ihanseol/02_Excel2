Private Sub CommandButton1_Click()
    If Workbooks.Count = 1 Then
        MsgBox "Please Open 기사용관정, 파일 ", vbOKOnly
        Exit Sub
    End If

   WB_NAME = GetOtherFileName
   
   ' Call mDeleteAllActiveXButtons(WB_NAME)
   
   
   Call DeleteAllActiveXControls(WB_NAME)
   Call SaveJustXLSX(WB_NAME)

End Sub


Private Sub CommandButton2_Click()
    If Workbooks.Count = 1 Then
        MsgBox "Please Open 기사용관정, 파일 ", vbOKOnly
        Exit Sub
    End If

   WB_NAME = GetOtherFileName
   Call DeleteHiddenSheets(WB_NAME)

End Sub





