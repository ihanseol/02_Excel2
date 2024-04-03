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
        Dim length As Long
        length = GetWindowText(hwnd, title, Len(title))
        
        If IS_DEBUG Then
            Print #FileNum, Left(title, length) ' Only write the actual text to the file
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

Function Initialize_Driver() As selenium.ChromeDriver
    Dim driver As selenium.ChromeDriver

    Set driver = New selenium.ChromeDriver
    driver.SetCapability "debuggerAddress", "localhost:9222"
    
    
    Set Initialize_Driver = driver
    
    ' driver.Get ("https://www.youtube.com/")
    ' Application.Wait Now + TimeValue("0:00:05")
    ' driver.Quit

End Function


Function Get_ScreenData(driver As selenium.ChromeDriver) As Variant

    Dim YONGDO_ORIGIN, YONGDO As String
    Dim SIMDO, DIAMETER, HP, Q, TOCHOOL As String
    
    Dim rYongDoOrigin, rYongDo As String
    Dim rSimDo, rDiameter, rHp, rQ, rTochool As Single
    
    ' Define CSS selectors
    YONGDO_ORIGIN = "#tb_info_detail2 > tbody > tr:nth-child(3) > td:nth-child(4)"
    YONGDO = "#tb_info_detail2 > tbody > tr:nth-child(4) > td:nth-child(2)"
    SIMDO = "#tb_info_detail2 > tbody > tr:nth-child(5) > td:nth-child(2)"
    DIAMETER = "#tb_info_detail2 > tbody > tr:nth-child(5) > td:nth-child(4)"
    HP = "#tb_info_detail2 > tbody > tr:nth-child(7) > td:nth-child(4)"
    Q = "#tb_info_detail2 > tbody > tr:nth-child(7) > td:nth-child(2)"
    TOCHOOL = "#tb_info_detail2 > tbody > tr:nth-child(8) > td"
    
    ' Retrieve data
    rYongDoOrigin = driver.FindElementByCss(YONGDO_ORIGIN).Text
    rYongDo = driver.FindElementByCss(YONGDO).Text
    rSimDo = driver.FindElementByCss(SIMDO).Text
    rDiameter = driver.FindElementByCss(DIAMETER).Text
    rHp = driver.FindElementByCss(HP).Text
    rQ = driver.FindElementByCss(Q).Text
    rTochool = driver.FindElementByCss(TOCHOOL).Text
    
    ' Create an array to hold the data
    Dim data(1 To 7) As Variant
    data(1) = rYongDoOrigin
    data(2) = rYongDo
    data(3) = rSimDo
    data(4) = rDiameter
    data(5) = rHp
    data(6) = rQ
    data(7) = rTochool
    
    ' Return the array
    Get_ScreenData = data

End Function



Sub Main()
Dim driver As selenium.ChromeDriver
Dim Result() As Variant

Set driver = Initialize_Driver()
Result = Get_ScreenData(driver)

For i = LBound(Result) To UBound(Result)

    Debug.Print Result(i)
Next i

End Sub







Sub AllColect()

    Dim YONGDO_ORIGIN, YONGDO As String
    Dim SIMDO, DIAMETER, HP, Q, TOCHOOL As String
    Dim i As Integer
    
    Dim rYongDoOrigin, rYongDo As String
    Dim rSimDo, rDiameter, rHp, rQ, rTochool As Single
    Dim driver As selenium.ChromeDriver
    
    Dim hwnd As LongPtr
    
    hwnd = GetWindowHandle("chrome")
     
'    Set driver = New selenium.ChromeDriver
'    driver.SetCapability "debuggerAddress", " localhost:9222/devtools/browser/" & hwnd
    
    
    Dim selenium As New SeleniumBasic
    Dim chrome_options As WebDriverOptions
    Set chrome_options = selenium.ChromeOptions
    
    chrome_options.AddArgument "headless"
    
    Dim driver As WebDriver
    Set driver = selenium.ChromeDriver(chrome_options)
        

    ' Define CSS selectors
    YONGDO_ORIGIN = "#tb_info_detail2 > tbody > tr:nth-child(3) > td:nth-child(4)"
    YONGDO = "#tb_info_detail2 > tbody > tr:nth-child(4) > td:nth-child(2)"
    SIMDO = "#tb_info_detail2 > tbody > tr:nth-child(5) > td:nth-child(2)"
    DIAMETER = "#tb_info_detail2 > tbody > tr:nth-child(5) > td:nth-child(4)"
    HP = "#tb_info_detail2 > tbody > tr:nth-child(7) > td:nth-child(4)"
    Q = "#tb_info_detail2 > tbody > tr:nth-child(7) > td:nth-child(2)"
    TOCHOOL = "#tb_info_detail2 > tbody > tr:nth-child(8) > td"
    
    ' Retrieve data
    rYongDoOrigin = driver.FindElementByCss(YONGDO_ORIGIN).Text
    rYongDo = driver.FindElementByCss(YONGDO).Text
    rSimDo = driver.FindElementByCss(SIMDO).Text
    rDiameter = driver.FindElementByCss(DIAMETER).Text
    rHp = driver.FindElementByCss(HP).Text
    rQ = driver.FindElementByCss(Q).Text
    rTochool = driver.FindElementByCss(TOCHOOL).Text
    
    ' Create an array to hold the data
    Dim data(1 To 7) As Variant
    data(1) = rYongDoOrigin
    data(2) = rYongDo
    data(3) = rSimDo
    data(4) = rDiameter
    data(5) = rHp
    data(6) = rQ
    data(7) = rTochool
    

    For i = 1 To 7
        Debug.Print data(i)
    Next i


End Sub
