Attribute VB_Name = "mod_SavePrn"
Global WB_NAME      As String

Public Function MyDocsPath() As String
    MyDocsPath = Environ$("USERPROFILE") & "\" & "Documents"
    Debug.Print MyDocsPath
End Function

Public Function WB_HEAD() As String
    Dim num As Integer
    
    num = GetNumbers(Worksheets("Input").Range("I54").Value)
    
    If num >= 10 Then
        WB_HEAD = MyDocsPath + "\" + Left(ThisWorkbook.name, 6)
    Else
        WB_HEAD = MyDocsPath + "\" + Left(ThisWorkbook.name, 5)
    End If
    
    Debug.Print WB_HEAD
    
End Function

Sub PrintSheetToPDF(ws As Worksheet, Optional filename As String = "None")
    Dim filePath As String
    
    
    If filename = "None" Then
        filePath = MyDocsPath & "\" & shInput.Range("I54").Value & ".pdf"
    Else
        filePath = MyDocsPath + "\" + filename
    End If
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           filename:=filePath, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=False ' Change to False if you don't want to open it automatically

    ' MsgBox "PDF saved at: " & filePath, vbInformation, "Success"
End Sub


Sub PrintSheetToPDF_Long(ws As Worksheet)
    Call PrintSheetToPDF(ws)
End Sub

Sub PrintSheetToPDF_LS(ws As Worksheet, filename As String)
    Call PrintSheetToPDF(ws, filename)
End Sub


Sub janggi_01()
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs filename:= _
                          WB_HEAD + "_janggi_01.dat", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
  
    Application.DisplayAlerts = True
  
End Sub

Sub janggi_02()
    
    Application.DisplayAlerts = False

    ActiveWorkbook.SaveAs filename:= _
                          WB_HEAD + "_janggi_02.dat", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
                          
   Application.DisplayAlerts = True
   
End Sub

Sub recover_01()
    Debug.Print WB_HEAD
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs filename:= _
                          WB_HEAD + "_recover_01.dat", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = True
End Sub

Sub step_01()
    Range("a1").Select
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs filename:= _
                          WB_HEAD + "_step_01.dat", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = True
    
End Sub

Sub save_original()

    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs filename:=WB_HEAD + "_OriginalSaveFile", FileFormat:= _
                          xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    
    Application.DisplayAlerts = True
    
End Sub







