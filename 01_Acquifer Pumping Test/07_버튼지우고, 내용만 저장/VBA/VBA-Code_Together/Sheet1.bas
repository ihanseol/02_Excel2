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


' Erase All ActiveX Control In Send
Private Sub CommandButton3_Click()


    Call OpenXLSMFilesInSend
  

End Sub


Sub OpenXLSMFilesInSend()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    
    folderPath = "d:\05_Send\"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Downloads folder not found!"
        Exit Sub
    End If
    
    fileName = Dir(folderPath & "*.xlsm")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        Call DeleteAllActiveXControls(fileName)
        Call SaveJustXLSX(fileName)
        
        wb.Close SaveChanges:=False
        fileName = Dir
    Loop
    
    MsgBox "All .xlsm files in Downloads folder have been opened."
End Sub



Sub OpenXLSMFilesInDownloads()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    
    folderPath = Environ("USERPROFILE") & "\Downloads\"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Downloads folder not found!"
        Exit Sub
    End If
    fileName = Dir(folderPath & "*.xlsm")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)

        fileName = Dir
    Loop
    
    MsgBox "All .xlsm files in Downloads folder have been opened."
End Sub


Sub OpenXlsmFiles()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim sFolder As String: sFolder = "C:\Users\YourUsername\Downloads"
    
    If Dir(sFolder) <> "" Then
        
        For Each fFile In fso.GetFolder(sFolder).Files
            If Right(fFile.Name, 4) = ".xlsm" Then
                
                Workbooks.Open (fFile.Path)
                Application.ScreenUpdating = True
            
            End If
        
        Next fFile
    
    Else
        MsgBox "The Downloads folder was not found. Please check the path.", vbCritical, "Error"
    
    End If

End Sub

