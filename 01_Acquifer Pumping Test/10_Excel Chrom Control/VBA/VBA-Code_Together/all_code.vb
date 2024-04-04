
'This Module is Empty 
'This Module is Empty 
Option Explicit

' Declare Windows API functions
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long

Dim hwnds As Collection

' Callback function to enumerate windows
Private Function EnumWindowsProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
    Dim windowText As String
    Dim Length As Long
    
    If IsWindowVisible(hwnd) Then
        Length = GetWindowTextLength(hwnd)
        If Length > 0 Then
            windowText = Space$(Length + 1)
            GetWindowText hwnd, windowText, Length + 1
            If InStr(1, windowText, "Google Chrome", vbTextCompare) > 0 Then
                hwnds.Add hwnd
            End If
        End If
    End If
    EnumWindowsProc = 1  ' Continue enumeration
End Function




Public Sub ListChromeWindows()
    Set hwnds = New Collection
    EnumWindows AddressOf EnumWindowsProc, 0
    
    ' Output the handles
    Dim i As Integer
    For i = 1 To hwnds.Count
        Debug.Print hwnds(i)
    Next i
End Sub

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



Sub ExamGetWellData()
    Dim chromePath As String
    Dim ChromeOptions As New selenium.ChromeOptions
    Dim driver As New selenium.ChromeDriver
    
    chromePath = "C:\Path\To\chromedriver.exe" ' Update with the path to your chromedriver.exe
    
    ' Set Chrome options
    ChromeOptions.AddArgument "remote-debugging-port=9222"
    
    ' Start Chrome driver with options
    driver.Start "chrome", chromePath, ChromeOptions
    
    ' Navigate to the webpage
    driver.Get "about:blank" ' Update with the URL of the webpage you want to navigate to
    
    ' Wait for the page to load (you may need to implement a wait condition)
    Application.Wait Now + TimeValue("0:00:05") ' Wait for 5 seconds (adjust as needed)
    
    ' Find the elements and interact with the webpage here
    
    ' Close the browser
    driver.Quit
End Sub

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
Function GetChrome(ByRef uia As CUIAutomation) As IUIAutomationElement
    Dim el_Desktop As IUIAutomationElement
    Set el_Desktop = uia.GetRootElement
    
    Dim el_ChromeWins As IUIAutomationElementArray
    Dim el_ChromeWin As IUIAutomationElement
    Dim cnd_ChromeWin As IUIAutomationCondition

    ' check the window with class name as "Chrome_WidgetWin_1"
    Set cnd_ChromeWin = uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Chrome_WidgetWin_1")
    Set el_ChromeWins = el_Desktop.FindAll(TreeScope_Children, cnd_ChromeWin)
    Set el_ChromeWin = Nothing
    
    If el_ChromeWins.Length = 0 Then
        Debug.Print """Chrome_WidgetWin_1"" not found"
        Exit Function
    End If
    
    Dim count_ChromeWins As Integer
    Dim CurWinTitle As String ' Declare CurWinTitle variable
    For count_ChromeWins = 0 To el_ChromeWins.Length - 1
        CurWinTitle = el_ChromeWins.GetElement(count_ChromeWins).CurrentName
        If (InStr(1, CurWinTitle, "Chrome")) Then
            Set el_ChromeWin = el_ChromeWins.GetElement(count_ChromeWins)
            Exit For
        End If
    Next
    
    If el_ChromeWin Is Nothing Then
        Debug.Print "No Chrome Window Found"
        Exit Function
    End If
    
    Set GetChrome = el_ChromeWin
End Function


Function Chrome_GetCurrentURL()

    Dim strURL As String
    Dim uia As New CUIAutomation
    Dim el_ChromeWin As IUIAutomationElement
    Set el_ChromeWin = GetChrome(uia)
    
    If el_ChromeWin Is Nothing Then
        Debug.Print "Chrome does NOT exist" ' Corrected the typo in "doe NOT"
        Exit Function
    End If
    
    Dim cnd As IUIAutomationCondition
    Set cnd = uia.CreatePropertyCondition(UIA_NamePropertyId, "ò¢ò£ûúâ¤ßã?")
    
    Dim AddressBar As IUIAutomationElement
    Set AddressBar = el_ChromeWin.FindFirst(TreeScope_Subtree, cnd)
    
    If Not AddressBar Is Nothing Then ' Added a check to ensure AddressBar is not null
        strURL = AddressBar.GetCurrentPropertyValue(UIA_ValueValuePropertyId)
    Else
        Debug.Print "AddressBar not found"
    End If
    
    Chrome_GetCurrentURL = strURL
    
End Function


Sub test()
    Dim url As String
    url = Chrome_GetCurrentURL()
    Debug.Print url
End Sub


Option Explicit

' https://www.mrexcel.com/board/threads/vba-any-way-to-cycle-between-excel-window-and-browser-window.1108521/

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As LongPtr, ByVal wFlag As Long) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Boolean


' Public Const MyGlobalBoolean As Boolean = True

Public Const IS_DEBUG As Boolean = False

Sub Start()

    Application.OnTime Now, "'DisplayWindows " & True & "'"

End Sub



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

Sub DisplayWindows(Optional ByVal MaximizedState As Boolean = True)
     Dim hwnd As LongPtr
   
    
    Dim oActiveWindow  As Window, oWnd As Window
    
    ActiveWindow.Activate
    Set oActiveWindow = ActiveWindow

    For Each oWnd In ThisWorkbook.Windows
        oWnd.Activate
        If MaximizedState Then oWnd.WindowState = xlMaximized
        Delay 2
    Next oWnd
    
    hwnd = GetWindowHandle("Chrome")
    If hwnd Then
        Call BringWindowToFront(hwnd, MaximizedState)
        Delay 2
    End If
    
    VBA.AppActivate Application.Caption
    oActiveWindow.Activate
    
    MsgBox "Back to initial window." & vbCrLf & vbCrLf & "Done!", vbInformation
    
End Sub


Private Sub Delay(DelayTime As Single)

    Dim sngTimer As Single
    
    sngTimer = Timer
    Do
        DoEvents
    Loop Until Timer - sngTimer >= DelayTime

End Sub


Private Sub BringWindowToFront(ByVal hwnd As LongPtr, Optional ByVal MaximizedState As Boolean = True)
 
    Const SW_SHOW = 5
    Const SW_SHOWMAXIMIZED = 3
    Const SW_RESTORE = 9
    Const GW_OWNER = 4
    
    Dim lThreadID1 As Long, lThreadID2 As Long
    
    On Error Resume Next
    
    If hwnd <> GetForegroundWindow() Then
        lThreadID1 = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
        lThreadID2 = GetWindowThreadProcessId(hwnd, ByVal 0&)
        Call AttachThreadInput(lThreadID1, lThreadID2, True)
        Call SetForegroundWindow(hwnd)
        If IsIconic(GetNextWindow(hwnd, 4)) Then
            Call ShowWindow(GetNextWindow(hwnd, GW_OWNER), IIf(MaximizedState, SW_SHOWMAXIMIZED, SW_RESTORE))
        Else
            Call ShowWindow(GetNextWindow(hwnd, GW_OWNER), IIf(MaximizedState, SW_SHOWMAXIMIZED, SW_SHOW))
        End If
        DoEvents
        Call AttachThreadInput(lThreadID2, lThreadID1, True)
    End If
End Sub



Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    Dim fName As String
    
    Dim pathToExport As String
    pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        
        
        If item.CodeModule.CountOfLines <> 0 Then
            lineToPrint = item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        Else
            lineToPrint = "'This Module is Empty "
        End If
        
        
        fName = item.CodeModule.Name
        Debug.Print lineToPrint
        SaveTextToFile lineToPrint, pathToExport & fName & ".bas"
        
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub
Sub ConvertJsonToDictionary()
    Dim jsonStr As String
    Dim dict As Object
    
    ' Your JSON string
    jsonStr = "{""Name"": ""John"", ""Age"": 30, ""City"": ""New York""}"
    
    ' Convert JSON string to dictionary
    Set dict = JsonToDictionary(jsonStr)
    
    ' Print values from dictionary
    Debug.Print "Name: " & dict("Name")
    Debug.Print "Age: " & dict("Age")
    Debug.Print "City: " & dict("City")
End Sub

Function JsonToDictionary(jsonStr As String) As Object
    Dim scriptControl As Object
    Dim JsonObject As Object
    
    ' Create a ScriptControl object
    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
    
    ' Evaluate the JSON string to create a JSON object
    Set JsonObject = scriptControl.Eval("(" + jsonStr + ")")
    
    ' Convert JSON object to dictionary
    Set JsonToDictionary = JsonToDictionaryRecursive(JsonObject)
End Function

Function JsonToDictionaryRecursive(JsonObject As Object) As Object
    Dim dict As Object
    Dim Key As Variant
    
    ' Create a dictionary object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Iterate through the JSON object
    For Each Key In JsonObject.Keys
        If IsObject(JsonObject(Key)) Then
            ' Recursively convert nested JSON objects
            dict.Add Key, JsonToDictionaryRecursive(JsonObject(Key))
        Else
            ' Add key-value pairs to the dictionary
            dict.Add Key, JsonObject(Key)
        End If
    Next Key
    
    ' Return the dictionary
    Set JsonToDictionaryRecursive = dict
End Function

Option Explicit

Private ScriptEngine As scriptControl

Public Sub InitScriptEngine()
On Error Resume Next
    Set ScriptEngine = New scriptControl
    ScriptEngine.Language = "JScript"
On Error GoTo 0

    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub

Public Function DecodeJsonString(ByVal JSonString As String)
    Set DecodeJsonString = ScriptEngine.Eval("(" + JSonString + ")")
End Function

Public Function GetProperty(ByVal JsonObject As Object, ByVal propertyName As String) 'As Variant
    GetProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)
End Function

Public Function GetObjectProperty(ByVal JsonObject As Object, ByVal propertyName As String) 'As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)
End Function

Public Function GetKeys(ByVal JsonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim Key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", JsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function

Sub ConvertJsonToDictionary()
    Dim jsonStr As String
    Dim dict As Object
    
    ' Your JSON string
    jsonStr = "{""Name"": ""John"", ""Age"": 30, ""City"": ""New York""}"
    
    Call InitScriptEngine
    Debug.Print DecodeJsonString(jsonStr)
    
    
    
End Sub

