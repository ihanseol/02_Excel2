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

