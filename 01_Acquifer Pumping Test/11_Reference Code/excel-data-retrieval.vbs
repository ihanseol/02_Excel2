Option Explicit

' Enum for different retrieval methods
Private Enum RetrievalMethod
    SingleCell
    RangeRead
    WorksheetRead
    ADOConnection
End Enum

' Function to retrieve data from Excel file
Public Function RetrieveExcelData(ByVal filePath As String, Optional method As RetrievalMethod = RangeRead) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim result As Variant
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    ' Disable screen updating and calculations for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Different retrieval methods
    Select Case method
        Case SingleCell
            ' Individual cell retrieval (slowest)
            result = RetrieveSingleCellData(filePath)
        
        Case RangeRead
            ' Optimized range reading
            result = RetrieveRangeData(filePath)
        
        Case WorksheetRead
            ' Read entire worksheet
            result = RetrieveWorksheetData(filePath)
        
        Case ADOConnection
            ' ADO Connection method (fastest for large files)
            result = RetrieveADOData(filePath)
    End Select
    
    ' Restore Excel settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    RetrieveExcelData = result
    Exit Function
    
ErrorHandler:
    ' Error handling
    MsgBox "Error retrieving data: " & Err.Description
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Function

' Method 1: Single Cell Retrieval
Private Function RetrieveSingleCellData(ByVal filePath As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Open workbook without displaying alerts
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(1)  ' First worksheet
    
    ' Example of retrieving specific cells
    Dim result(1 To 5) As Variant
    result(1) = ws.Range("A1").Value
    result(2) = ws.Range("B2").Value
    result(3) = ws.Range("C3").Value
    result(4) = ws.Range("D4").Value
    result(5) = ws.Range("E5").Value
    
    ' Close workbook
    wb.Close SaveChanges:=False
    
    RetrieveSingleCellData = result
End Function

' Method 2: Range Reading
Private Function RetrieveRangeData(ByVal filePath As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Open workbook without displaying alerts
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(1)  ' First worksheet
    
    ' Read entire range at once
    Dim dataRange As Range
    Set dataRange = ws.Range("A1:E10")  ' Adjust range as needed
    
    ' Convert range to array for faster processing
    Dim result As Variant
    result = dataRange.Value
    
    ' Close workbook
    wb.Close SaveChanges:=False
    
    RetrieveRangeData = result
End Function

' Method 3: Entire Worksheet Reading
Private Function RetrieveWorksheetData(ByVal filePath As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' Open workbook without displaying alerts
    Set wb = Workbooks.Open(filePath, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(1)  ' First worksheet
    
    ' Find last used row and column
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Read entire worksheet data
    Dim result As Variant
    result = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
    
    ' Close workbook
    wb.Close SaveChanges:=False
    
    RetrieveWorksheetData = result
End Function

' Method 4: ADO Connection (Fastest for large files)
Private Function RetrieveADOData(ByVal filePath As String) As Variant
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim result As Variant
    
    ' Create connection string
    Dim connString As String
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                 "Data Source=" & filePath & ";" & _
                 "Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    
    ' Create and open connection
    Set conn = New ADODB.Connection
    conn.Open connString
    
    ' Execute query
    Set rs = conn.Execute("SELECT * FROM [Sheet1$]")
    
    ' Convert recordset to array
    result = rs.GetRows()
    
    ' Close and clean up
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    RetrieveADOData = result
End Function

' Example usage
Sub TestDataRetrieval()
    Dim filePath As String
    Dim result As Variant
    
    filePath = "C:\YourPath\YourFile.xlsx"
    
    ' Retrieve data using different methods
    result = RetrieveExcelData(filePath, RangeRead)
    
    ' Process or display result
    Debug.Print result(1, 1)  ' Print first cell
End Sub
