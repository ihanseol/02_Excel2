
Private Sub CommandButton_Print_Long_Click()
    Dim well As Integer
    well = GetNumbers(shInput.Range("I54").Value)

    Sheets("장회").Visible = True
    Sheets("장회").Activate
    Call PrintSheetToPDF_Long(Sheets("장회"), "w" + CStr(well))
    Sheets("장회").Visible = False
    
End Sub

Private Sub CommandButton_Print_LS_Click()
    Dim well As Integer
    
    Sheets("장회").Visible = True
    Sheets("단계").Visible = True
    well = GetNumbers(shInput.Range("I54").Value)
    
    Sheets("단계").Activate
    Call PrintSheetToPDF_LS(Sheets("단계"), "w" + CStr(well) + "-1.pdf")
    Sheets("단계").Visible = False
    
    Sheets("장회").Activate
    Call PrintSheetToPDF_LS(Sheets("장회"), "w" + CStr(well) + "-2.pdf")
    Sheets("장회").Visible = False
    
End Sub



Private Sub CommandButton1_Click()
    Call recover_01
End Sub

Private Sub CommandButton2_Click()
    Call save_original
End Sub

Private Sub CommandButton3_Click()

Sheets("장회").Visible = True
Sheets("장회14").Visible = True
Sheets("단계").Visible = True
Sheets("장기28").Visible = True
Sheets("장기14").Visible = True
Sheets("회복").Visible = True
Sheets("회복12").Visible = True

End Sub

Private Sub CommandButton4_Click()

Sheets("장회").Visible = False
Sheets("장회14").Visible = False
Sheets("단계").Visible = False
Sheets("장기28").Visible = False
Sheets("장기14").Visible = False
Sheets("회복").Visible = False
Sheets("회복12").Visible = False

End Sub

Private Sub Worksheet_Activate()
    Dim gong1, gong2 As String
    Dim gong As Long
    Dim er As Integer
    Dim cellformula As String
    

'    gong = Val(CleanString(shInput.Range("J48").Value))
'
'    gong1 = "W-" & CStr(gong)
'    gong2 = shInput.Range("i54").Value
'
'    If gong1 <> gong2 Then
'        'MsgBox "different : " & g1 & " g2 : " & g2
'        shInput.Range("i54").Value = gong1
'    End If
    

    er = GetEffectiveRadius
        
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!K8"
        
        Case erRE2
            cellformula = "=SkinFactor!K9"
            
        Case erRE3
            cellformula = "=SkinFactor!K10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select
    
    Range("A28").Formula = cellformula
    
End Sub


