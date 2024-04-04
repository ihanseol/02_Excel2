Option Explicit


'in here excel find chrome window handle
'and use that handle find previous session


Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Boolean


' Public Const MyGlobalBoolean As Boolean = True

Public Const IS_DEBUG As Boolean = False


Sub DeleteFileIfExists(filePath As String)
    On Error Resume Next
    Kill filePath ' Attempt to delete the file
    On Error GoTo 0
    If Len(Dir(filePath)) > 0 Then
        MsgBox "Unable to delete file '" & filePath & "'. The file may be in use.", vbExclamation
    End If
End Sub

Public Function GetWindowHandle(ByVal substring As String) As LongPtr
    Dim hwnd As LongPtr
    Dim FileNum As Integer
    Dim filePath As String

    hwnd = FindWindowEx(0&, 0&, vbNullString, vbNullString)
    
    substring = LCase(substring)
    
    If hwnd = 0 Then
        GetWindowHandle = "No windows found."
        Exit Function
    End If

    If IS_DEBUG Then
        filePath = ThisWorkbook.Path & "\" & ActiveSheet.Name & ".csv"
        DeleteFileIfExists filePath
    
        FileNum = FreeFile
        Open filePath For Output As FileNum
    End If

    Do While hwnd <> 0
        Dim title As String * 255
        Dim Length As Long
        Length = GetWindowText(hwnd, title, Len(title))
        
        If IS_DEBUG Then
            Print #FileNum, Left(title, Length) ' Only write the actual text to the file
        End If
        ' Check if the window title contains the specified substring
        If InStr(1, title, substring, vbTextCompare) > 0 Then
            GetWindowHandle = hwnd
            
            If IS_DEBUG Then
                Close FileNum
             End If
            Exit Function
        End If
               
        
        hwnd = FindWindowEx(0&, hwnd, vbNullString, vbNullString)
    Loop
    
    GetWindowHandle = 0&
    
    If IS_DEBUG Then
        Close FileNum
    End If
    
End Function

Sub TestFindWindowHandle_aa()
   Dim hwnd As LongPtr
    
    hwnd = GetWindowHandle("chrome")
    
    If hwnd Then
        MsgBox hwnd
        Call GetWellData(hwnd)
    Else
        MsgBox "Did not found ...."
    End If

End Sub



Sub GetWellData(ByVal hwnd As LongPtr)
    Dim chromePath As String

    Dim driver As New ChromeDriver
    Dim ddl As selenium.SelectElement
    
    
    chromePath = "localhost:9222/devtools/browser/" & hwnd
    Debug.Print chromePath
    
    
     'chrome_options.add_experimental_option('debuggerAddress', f"localhost:9222/devtools/browser/{window_handle}")
    
    Dim chromeBrowser As selenium.ChromeDriver

      Set chromeBrowser = New selenium.ChromeDriver
      chromeBrowser.SetCapability "debuggerAddress", "localhost:9222"
      chromeBrowser.Get ("https://www.google.com")


'    driver.Start "chrome", chromePath
'    driver.AddArgument "remote-debugging-port=9222"
'    driver.Window.Maximize
'    driver.Get "https://www.daum.net/"
    
    Application.Wait Now + TimeValue("0:00:05")


    driver.Quit
End Sub



