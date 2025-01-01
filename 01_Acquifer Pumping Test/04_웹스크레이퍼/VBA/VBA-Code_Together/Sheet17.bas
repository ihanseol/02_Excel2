Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
