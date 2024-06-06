Private Sub Workbook_Open()
    Call InitialSetColorValue
    Sheets("Well").SingleColor.Value = True
    Sheets("Recharge").cbCheSoo.Value = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
 ' Call InitialSetColorValue
End Sub


