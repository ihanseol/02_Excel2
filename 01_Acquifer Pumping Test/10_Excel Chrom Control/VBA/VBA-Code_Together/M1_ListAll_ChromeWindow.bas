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
    Dim length As Long
    
    If IsWindowVisible(hwnd) Then
        length = GetWindowTextLength(hwnd)
        If length > 0 Then
            windowText = Space$(length + 1)
            GetWindowText hwnd, windowText, length + 1
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

