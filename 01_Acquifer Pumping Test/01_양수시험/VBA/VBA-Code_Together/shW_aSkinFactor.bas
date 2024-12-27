

Private Sub CommandButton_ExRE1_Click()
    Range("EffectiveRadius").Value = "경험식 1번"
    Range("D4").Value = Range("D5").Value

End Sub

Private Sub CommandButton_ExRE3_Click()
    Range("EffectiveRadius").Value = "경험식 3번"
    Range("D4").Value = Range("D5").Value
End Sub

Private Sub CommandButton_GetStepT_Click()
    Range("D4").Value = shW_StepTEST.Range("T4").Value
End Sub

Private Sub CommandButton_SkinFactor_Click()
    Range("EffectiveRadius").Value = "SkinFactor"
End Sub

Private Sub CommandButton1_Click()
    Call show_gachae
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
