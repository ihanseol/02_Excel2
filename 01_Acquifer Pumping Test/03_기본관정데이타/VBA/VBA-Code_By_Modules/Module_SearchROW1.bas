Attribute VB_Name = "Module_SearchROW1"


Function CellContains(searchRange As Range, searchValue As String) As Boolean
    CellContains = InStr(1, LCase(searchRange.value), LCase(searchValue)) > 0
End Function


Function FindCellByLoopingPartialMatch() As String

    Dim ws As Worksheet
    Dim cell As Range
    Dim address As String
     
     For Each cell In Range("A1:AZ1").Cells
        Debug.Print cell.address, cell.value
    
        If CellContains(cell, "") Then
            address = cell.address
            Exit For
        End If
    Next
    FindCellByLoopingPartialMatch = address
    
End Function

Sub test()
    Dim rg As String
    rg = FindCellByLoopingPartialMatch
    Debug.Print "the result: ", rg, Range(rg).value
End Sub

