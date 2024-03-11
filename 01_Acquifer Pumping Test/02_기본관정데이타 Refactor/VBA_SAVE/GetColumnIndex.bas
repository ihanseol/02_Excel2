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
