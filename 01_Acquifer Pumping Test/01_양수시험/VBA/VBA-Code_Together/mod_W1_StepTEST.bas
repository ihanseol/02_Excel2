Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


Sub CutDownNumber(po As String, cutdown As Integer)
    Dim i, chrcode As Integer
    For i = 1 To 5
        Cells(i + 43, po).Value = Format(Round(Cells(i + 43, po).Value, cutdown), "###0.000")
    Next i
End Sub

Sub EraseCellData(ByVal str_range As String)
    With Range(str_range)
        .Value = ""
    End With
End Sub



Sub WriteStringEtc()
    Dim i As Integer
    Dim cv1, cv2, cv3, append As String
    Dim arr() As Variant
    arr = Array(0, 120, 240, 360, 480)
    
    For i = 1 To 5
        If i = 5 Then
            append = ""
        Else
            append = vbLf
        End If
        
        cv1 = cv1 & CStr(i) & append
        cv2 = cv2 & CStr(arr(i - 1)) & append
        cv3 = cv3 & CStr(120) & append
    Next i
    
    Cells(64, "v").Value = cv1
    Cells(64, "w").Value = cv2
    Cells(64, "x").Value = cv3
End Sub

Function ConcatenateCells(inRange As String) As String
    Dim cell As Range
    Dim concatenatedValue As String
    Dim sFormat(1 To 5) As String
    Dim i As Integer
    
    
    sFormat(1) = "###0"
    sFormat(2) = "###0.00"
    sFormat(3) = "###0.00"
    sFormat(4) = "###0.000"
    sFormat(5) = "###0.000000"
    
    i = Asc(Left(inRange, 1)) - Asc("P")
        
    For Each cell In Range(inRange)
        concatenatedValue = concatenatedValue & Format(cell.Value, sFormat(i)) & vbLf
    Next cell
    
     ConcatenateCells = Left(concatenatedValue, Len(concatenatedValue) - 1)
End Function


Function get_chart_equation(ByVal chartname) As String
    Dim objTrendline As Trendline
    Dim strEquation As String
    
    With ActiveSheet.ChartObjects(chartname).Chart
        Set objTrendline = .SeriesCollection(1).Trendlines(1)
        With objTrendline
            .DisplayRSquared = False
            .DisplayEquation = True
            strEquation = .DataLabel.Text
        End With
    End With
    
    get_chart_equation = strEquation
End Function

Function split_string(ByVal name As String) As String()
    Dim myarray()   As String
    
    myarray = Split(name)
    split_string = myarray
End Function

Sub get_chart7(ByRef c1 As Double, ByRef d1 As Double)
    Dim eq          As String
    Dim t_array()   As String
    Dim c, d        As Double
    
    eq = get_chart_equation("Chart 7")
    t_array = split_string(eq)
    
    c = CDbl(t_array(2))
    d = CDbl(t_array(5))
    
    Range("p37").Value = c
    Range("p38").Value = d
    
    c1 = c
    d1 = d
End Sub

Sub get_chart8(ByRef c1 As Double, ByRef d1 As Double)
    Dim eq          As String
    Dim t_array()   As String
    Dim c, d        As Double
    
    eq = get_chart_equation("Chart 8")
    t_array = split_string(eq)
    
    c = Abs(Round(CDbl(t_array(2)), 3))
    d = Round(CDbl(t_array(5)), 3)
    
    Range("q37").Value = c
    Range("q38").Value = d
    
    c1 = c
    d1 = d
End Sub

Sub ChangeCharts()
    Dim myChart     As ChartObject
    
    For Each myChart In ActiveSheet.ChartObjects
        myChart.Chart.Refresh
    Next myChart
End Sub


Sub set_CB_ALL()

    Call set_CB1
    MsgBox "SetCB1 Complete and Next setCB2 .... ", vbOKOnly
    Call set_CB2
    
End Sub


Sub set_CB1()
    Dim c           As Double
    Dim d           As Double
    
    On Error GoTo ErrorCheck
    Call get_chart7(c, d)
    
    Range("a31").Value = c
    Range("b31").Value = d
    Exit Sub
    
ErrorCheck:
    ' MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Sub set_CB2()
    Dim c           As Double
    Dim d           As Double
    
    On Error GoTo ErrorCheck
    Call get_chart8(c, d)
    
    Range("b38").Value = c
    Range("c38").Value = d
    Range("a38").Value = Range("d39").Value
    Exit Sub
    
ErrorCheck:
    ' MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

