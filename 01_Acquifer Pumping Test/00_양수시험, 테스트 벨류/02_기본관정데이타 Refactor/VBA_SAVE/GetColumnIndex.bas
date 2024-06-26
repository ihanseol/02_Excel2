Function GetColumnIndex(ByVal columnName As String) As Integer
    ' Define array to store column indices
    Dim columnIndices As Variant
    columnIndices = Array( _
        11, 13, 2, 3, 7, 8, 9, 10, _
        32, 33, 4, 5, 6, 12, 14, _
        35, 36, 37, 15, 16, 17, 18, _
        19, 20, 21, 22, 23, 24, 25, _
        26, 38, 39, 40, 27, 28, 30, _
        31, 29, 34 _
    )

    ' Define array to store column names
    Dim columnNames As Variant
    columnNames = Array( _
        "Q", "hp", "natural", "stable", "radius", "Rw", "well_depth", "casing", _
        "C", "B", "recover", "Sw", "delta_h", "delta_s", "daeSoo", _
        "T0", "S0", "ER_MODE", "T1", "T2", "TA", "S1", _
        "S2", "K", "time_", "shultze", "webber", "jacob", "skin", _
        "er", "ER1", "ER2", "ER3", "qh", "qg", "sd1", _
        "sd2", "q1", "ratio" _
    )

    ' Find index of columnName in columnNames array
    Dim index As Integer
    index = Application.match(columnName, columnNames, 0)

    ' Check if columnName exists in columnNames array
    If IsNumeric(index) Then
        GetColumnIndex = columnIndices(index - 1)
    Else
        ' Return -1 if columnName is not found
        GetColumnIndex = -1
    End If
End Function

'***************************************************************************************************


Function GetColumnIndex(ByVal columnName As String) As Integer
    Dim columnIndexMap As Object
    Set columnIndexMap = CreateObject("Scripting.Dictionary")

    ' Define column name to index mappings
    With columnIndexMap
        .Add "Q", 11
        .Add "hp", 13
        .Add "natural", 2
        .Add "stable", 3
        .Add "radius", 7
        .Add "Rw", 8
        .Add "well_depth", 9
        .Add "casing", 10
        .Add "C", 32
        .Add "B", 33
        .Add "recover", 4
        .Add "Sw", 5
        .Add "delta_h", 6
        .Add "delta_s", 12
        .Add "daeSoo", 14
        .Add "T0", 35
        .Add "S0", 36
        .Add "ER_MODE", 37
        .Add "T1", 15
        .Add "T2", 16
        .Add "TA", 17
        .Add "S1", 18
        .Add "S2", 19
        .Add "K", 20
        .Add "time_", 21
        .Add "shultze", 22
        .Add "webber", 23
        .Add "jacob", 24
        .Add "skin", 25
        .Add "er", 26
        .Add "ER1", 38
        .Add "ER2", 39
        .Add "ER3", 40
        .Add "qh", 27
        .Add "qg", 28
        .Add "sd1", 30
        .Add "sd2", 31
        .Add "q1", 29
        .Add "ratio", 34
    End With

    ' Check if columnName exists in the dictionary
    If columnIndexMap.Exists(columnName) Then
        GetColumnIndex = columnIndexMap(columnName)
    Else
        ' Return -1 if columnName is not found
        GetColumnIndex = -1
    End If
End Function

'***************************************************************************************************


Function GetColumnIndex(ByVal columnName As String) As Integer
    Dim colIndex As Integer

    Select Case columnName
        Case "Q"
            colIndex = 11
        Case "hp"
            colIndex = 13
        
        
        Case "natural"
            colIndex = 2
        Case "stable"
            colIndex = 3
        Case "radius"
            colIndex = 7
        Case "Rw"
            colIndex = 8
        
        Case "well_depth"
            colIndex = 9
        Case "casing"
           colIndex = 10
        
        Case "C"
           colIndex = 32
         Case "B"
            colIndex = 33
        
        
        Case "recover"
            colIndex = 4
        Case "Sw"
            colIndex = 5
        
        Case "delta_h"
            colIndex = 6
        Case "delta_s"
            colIndex = 12
    
        Case "daeSoo"
           colIndex = 14
            
  '--------------------------------------------------------------
  
       Case "T0"
           colIndex = 35
        Case "S0"
           colIndex = 36
       Case "ER_MODE"
           colIndex = 37
                  
        Case "T1"
           colIndex = 15
        Case "T2"
            colIndex = 16
        Case "TA"
           colIndex = 17
            
       Case "S1"
           colIndex = 18
        Case "S2"
            colIndex = 19
        
        Case "K"
           colIndex = 20
        Case "time_"
            colIndex = 21
            
        Case "shultze"
           colIndex = 22
        Case "webber"
            colIndex = 23
        Case "jacob"
            colIndex = 24
                    
                        
       Case "skin"
            colIndex = 25
        Case "er"
            colIndex = 26
            
        Case "ER1"
            colIndex = 38
        Case "ER2"
            colIndex = 39
        Case "ER3"
            colIndex = 40

        Case "qh"
            colIndex = 27
        Case "qg"
            colIndex = 28
            
        Case "sd1"
            colIndex = 30
        Case "sd2"
            colIndex = 31
        Case "q1"
            colIndex = 29
        Case "ratio"
            colIndex = 34
    End Select

    GetColumnIndex = colIndex
End Function



'***************************************************************************************************
