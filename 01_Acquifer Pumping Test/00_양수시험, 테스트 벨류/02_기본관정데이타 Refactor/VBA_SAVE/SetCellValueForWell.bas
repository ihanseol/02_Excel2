Sub SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim wellData As Variant
    Dim numberFormats As Variant

    ' Store number formats for each dataArrayName
    numberFormats = Array("", "", "0.00", "0.00", "", "0.0000000", "0.0000", "0.0000", "", "", _
                         "0.00", "0.00", "0.00", "0.00", "0.0000", "0.0000", "0.00", "0.00", _
                         "0.00", "0.00", "0.00", "0.00", "0.00", "0.00", "0.0000", "0.0000", _
                         "0.0000", "0.00", "0.00", "0.00", "0.0000", "0.0000", "0.0%", "0.0000", "0.0000")

    ' Get value from dataCell
    wellData = dataCell.Value
    
    ' Set value and number format based on dataArrayName
    With Cells(4 + wellIndex, GetColumnIndex(dataArrayName))
        .Value = wellData
        .NumberFormat = numberFormats(GetColumnIndex(dataArrayName) - 1)
    End With
End Sub



'**********************************************************************************'




Sub SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim wellData As Variant
    Dim numberFormats As Object
    Set numberFormats = CreateObject("Scripting.Dictionary")

    ' Define number formats for each dataArrayName
    With numberFormats
        .Add "recover", "0.00"
        .Add "Sw", "0.00"
        .Add "S2", "0.0000000"
        .Add "T1", "0.0000"
        .Add "T2", "0.0000"
        .Add "TA", "0.0000"
        .Add "qh", "0."
        .Add "qg", "0.00"
        .Add "q1", "0.00"
        .Add "sd1", "0.00"
        .Add "sd2", "0.00"
        .Add "skin", "0.0000"
        .Add "er", "0.0000"
        .Add "ratio", "0.0%"
        .Add "T0", "0.0000"
        .Add "S0", "0.0000"
        .Add "delta_s", "0.00"
        .Add "time_", "0.00"
        .Add "shultze", "0.00"
        .Add "webber", "0.00"
        .Add "jacob", "0.00"
        
    End With

    ' Get value from dataCell
    wellData = dataCell.value
    
    Cells(4 + wellIndex, 1).value = "W-" & wellIndex
    
    ' Set value and number format based on dataArrayName
    With Cells(4 + wellIndex, GetColumnIndex(dataArrayName))
        .value = wellData
        If numberFormats.Exists(dataArrayName) Then
            .NumberFormat = numberFormats(dataArrayName)
        End If
    End With
End Sub



'**********************************************************************************'





Sub SetCellValueForWell(ByVal wellIndex As Integer, ByVal dataCell As Range, ByVal dataArrayName As String)
    Dim wellData As Variant

    wellData = dataCell.value
    
    
    Cells(4 + wellIndex, 1).value = "W-" & wellIndex
    Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).value = wellData
    
    If dataArrayName = "recover" Or dataArrayName = "Sw" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "S2" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000000"
    ElseIf dataArrayName = "T1" Or dataArrayName = "T2" Or dataArrayName = "TA" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    ElseIf dataArrayName = "qh" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0."
    ElseIf dataArrayName = "qg" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "q1" Or dataArrayName = "sd1" Or dataArrayName = "sd2" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.00"
    ElseIf dataArrayName = "skin" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).value = Format(wellData, "0.0000")
    ElseIf dataArrayName = "er" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    ElseIf dataArrayName = "ratio" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0%"
    ElseIf dataArrayName = "T0" Or dataArrayName = "S0" Then
        Cells(4 + wellIndex, GetColumnIndex(dataArrayName)).NumberFormat = "0.0000"
    End If
End Sub



'**********************************************************************************'

