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

Sub janggi_01()
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_janggi_01.dat", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
  
    Application.DisplayAlerts = True
  
End Sub

Sub janggi_02()
    
    Application.DisplayAlerts = False

    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_janggi_02.dat", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
                          
   Application.DisplayAlerts = True
   
End Sub

Sub recover_01()
    Debug.Print WB_HEAD
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_recover_01.dat", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = True
End Sub

Sub step_01()
    Range("a1").Select
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_step_01.dat", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
    
    Application.DisplayAlerts = True
    
End Sub

Sub save_original()

    Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs fileName:=WB_HEAD + "_OriginalSaveFile", FileFormat:= _
                          xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    
    Application.DisplayAlerts = True
    
End Sub







