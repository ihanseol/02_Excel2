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


