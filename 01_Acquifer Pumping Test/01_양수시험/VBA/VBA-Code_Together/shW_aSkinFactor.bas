Private Sub CommandButton_Experimential_Click()

    Range("EffectiveRadius").Value = "경험식 1번"
    Range("D4").Value = Range("D5").Value

End Sub

Private Sub CommandButton1_Click()
    Call show_gachae
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
