Attribute VB_Name = "modFrame"


Public Sub MakeFrameFromSelection()
Attribute MakeFrameFromSelection.VB_ProcData.VB_Invoke_Func = "d\n14"
' e30:i43

    Dim result As Variant
    Dim start_alpha, end_alpha As String
    Dim start_num, end_num, i  As Integer
    
    Dim rng_str As String
    
    rng_str = GetRangeStringFromSelection()
    rng_str = Replace(rng_str, "$", "")
    
    If Not CheckSubstring(rng_str, ":") Then
        Exit Sub
    End If

    Debug.Print rng_str
    Call MakeColorFrame(rng_str)
   
End Sub


Public Sub MakeColorFrame(ByVal str_rng As String)
' e30:i43

    Dim result As Variant
    Dim start_alpha, end_alpha As String
    Dim start_num, end_num, i  As Integer
    
    Range(str_rng).Select
    
    result = ExtractStringPartsUsingRegex(str_rng)
    start_alpha = result(0)
    end_alpha = result(1)
    start_num = result(2)
    end_num = result(3)
    
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    
    
    Range(start_alpha & start_num & ":" & end_alpha & start_num).Select
    
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    
    For i = start_num + 2 To end_num Step 2
    
        Range(start_alpha & i & ":" & end_alpha & i).Select
        
        With Selection.Interior
            .pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .themeColor = xlThemeColorDark1
            .TintAndShade = -4.99893185216834E-02
            .PatternTintAndShade = 0
        End With
    
    Next i
    
End Sub



Function ExtractStringPartsUsingRegex(inputString As String) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim result(3) As Variant
    
    ' Create a regex object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Define the pattern
    ' The pattern ([a-zA-Z]+)(\d+) matches any letter(s) followed by any digit(s)
    regex.pattern = "([a-zA-Z]+)(\d+)"
    regex.Global = True
    
    ' Execute the regex pattern on the input string
    Set matches = regex.Execute(inputString)
    
    ' Check if the pattern matched the expected number of parts
    If matches.count >= 2 Then
        ' Store the results in the result array
        result(0) = matches(0).SubMatches(0) ' Start Letter
        result(1) = matches(1).SubMatches(0) ' End Letter
        result(2) = CLng(matches(0).SubMatches(1)) ' Start Number
        result(3) = CLng(matches(1).SubMatches(1)) ' End Number
    Else
        ' Handle the case where the pattern did not match
        result(0) = ""
        result(1) = ""
        result(2) = 0
        result(3) = 0
    End If
    
    ' Return the result array
    ExtractStringPartsUsingRegex = result
End Function

Sub TestExtractStringParts()
    Dim inputString As String
    Dim result As Variant
    
    ' Given string
    inputString = "e30:i43"
    
    ' Call the function and get the result
    result = ExtractStringPartsUsingRegex(inputString)
    
    ' Output the results
    Debug.Print "Start Letter: " & result(0)
    Debug.Print "End Letter: " & result(1)
    Debug.Print "Start Number: " & result(2)
    Debug.Print "End Number: " & result(3)
End Sub


Sub test()
    Call MakeColorFrame("o22:s33")
End Sub

