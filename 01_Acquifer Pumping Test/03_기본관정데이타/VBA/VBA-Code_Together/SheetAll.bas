

Private Sub CommandButton_boryung_Click()
    Call importRainfall_button("BORYUNG")
End Sub

Private Sub CommandButton_buyeo_Click()
    Call importRainfall_button("BUYEO")
End Sub

Private Sub CommandButton_cheonan_Click()
    Call importRainfall_button("CHEONAN")
End Sub

Private Sub CommandButton_daejeon_Click()
    Call importRainfall_button("DAEJEON")
End Sub

Private Sub CommandButton_seosan_Click()
    Call importRainfall_button("SEOSAN")
End Sub


Private Sub CommandButton_Seoul_Click()
     Call importRainfall_button("SEOUL")
End Sub

Private Sub CommandButton1_Click()
    Call importRainfall
End Sub

Private Sub CommandButton2_Click()
    Range("b5:n34").ClearContents
End Sub



