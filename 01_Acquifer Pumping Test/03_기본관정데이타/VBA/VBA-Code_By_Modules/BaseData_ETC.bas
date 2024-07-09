Attribute VB_Name = "BaseData_ETC"
Option Explicit

'------------------------------------------------------------------------------------------
' 2022/6/11

Public Enum cellLowHi
    cellLOW = 0
    cellHI = 1
End Enum


Sub EraseCellData(ByVal str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Sub Delay(ByVal msg As String, ByVal n As Integer)
    Application.Wait (Now + TimeValue("0:00:" & n))
    MsgBox msg, vbOKOnly
End Sub


'Function GetNumberOfWell() As Integer
'    Dim save_name As String
'    Dim n As Integer
'
'    save_name = ActiveSheet.Name
'    Sheets("Well").Activate
'    Sheets("Well").Range("A30").Select
'    Selection.End(xlUp).Select
'    n = CInt(GetNumeric2(Selection.value))
'
'    GetNumberOfWell = n
'End Function


Function ColumnNumberToLetter(ByVal columnNumber As Integer) As String
    Dim dividend As Integer
    Dim modulo As Integer
    Dim columnName As String
    Dim result As String
    
    dividend = columnNumber
    result = ""
    
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnName = Chr(65 + modulo) & columnName
        dividend = (dividend - modulo) \ 26
    Loop
    
    ColumnNumberToLetter = columnName
End Function


Function ColumnLetterToNumber(ByVal columnLetter As String) As Long
    Dim i As Long
    Dim result As Long

    result = 0
    For i = 1 To Len(columnLetter)
        result = result * 26 + (Asc(UCase(Mid(columnLetter, i, 1))) - 64)
    Next i

    ColumnLetterToNumber = result
End Function


Sub BackGroundFill(rngLine As Range, FLAG As Boolean)

If FLAG Then
    rngLine.Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
Else
    rngLine.Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End If

End Sub

Function GetRowColumn(name As String) As Variant
    Dim acColumn, acRow As Variant
    Dim result(1 To 2) As Variant

    acColumn = Split(Range(name).Address, "$")(1)
    acRow = Split(Range(name).Address, "$")(2)

    '  Row = ActiveCell.Row
    '  col = ActiveCell.Column
    
    
    result(1) = acColumn
    result(2) = acRow

    Debug.Print acColumn, acRow
    GetRowColumn = result
End Function


' 이것은, Well 탭의 값을 가지고 검사하하는것이라서, 차이가 생긴다.
Function GetNumberOfWell() As Integer
    Dim save_name As String
    Dim n As Integer
    
    save_name = ActiveSheet.name
    With Sheets("Well")
        n = .Cells(.Rows.count, "A").End(xlUp).row
        n = CInt(GetNumeric2(.Cells(n, "A").value))
    End With
    
    GetNumberOfWell = n
End Function

Public Function sheets_count() As Long
    Dim i As Integer
    Dim nSheetsCount As Long
    Dim nWell As Long
    Dim strSheetsName() As String

    nSheetsCount = ThisWorkbook.Sheets.count
    nWell = 0

    ReDim strSheetsName(1 To nSheetsCount)

    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).name
        If ConvertToLongInteger(strSheetsName(i)) <> 0 Then
            nWell = nWell + 1
        End If
    Next i

    sheets_count = nWell
End Function


Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long

    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        If StrComp(ThisWorkbook.name, Workbook.name, vbTextCompare) = 0 Then
            GoTo NEXT_ITERATION
        End If
        
        If CheckSubstring(Workbook.name, "OriginalSaveFile") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    If Workbook Is Nothing Then
        GetOtherFileName = "Empty"
    Else
        GetOtherFileName = Workbook.name
    End If
    
End Function


Function CheckSubstring(ByVal str As String, ByVal chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function



Function ExtractNumberFromString(inputString As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "\d+"
    End With
    
    If regex.test(inputString) Then
        Set matches = regex.Execute(inputString)
        ExtractNumberFromString = matches(0)
    Else
        ExtractNumberFromString = "No numbers found"
    End If
End Function



Function GetNumeric2(ByVal CellRef As String) As String
    Dim StringLength, i  As Integer
    Dim result      As String
    
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
    Next i
    GetNumeric2 = result
End Function

'********************************************************************************************************************************************************************************
'Function Name                    : IsWorkBookOpen(ByVal OWB As String)
'Function Description             : Function to check whether specified workbook is open
'Data Parameters                  : OWB:- Specify name or path to the workbook. eg: "Book1.xlsx" or "C:\Users\Kannan.S\Desktop\Book1.xlsm"

'********************************************************************************************************************************************************************************
Function IsWorkBookOpen(ByVal OWB As String) As Boolean
    IsWorkBookOpen = False
    Dim wb          As Excel.Workbook
    Dim WBNAME      As String
    Dim WBPath      As String
    Dim OWBArray    As Variant
    
    Err.Clear
    
    On Error Resume Next
    OWBArray = Split(OWB, Application.PathSeparator)
    Set wb = Application.Workbooks(OWBArray(UBound(OWBArray)))
    WBNAME = OWBArray(UBound(OWBArray))
    WBPath = wb.Path & Application.PathSeparator & WBNAME
    
    If Not wb Is Nothing Then
        If UBound(OWBArray) > 0 Then
            If LCase(WBPath) = LCase(OWB) Then IsWorkBookOpen = True
        Else
            IsWorkBookOpen = True
        End If
    End If
    Err.Clear
    
End Function

'------------------------------------------------------------------------------------------

Public Function GetLengthByColor(ByVal tabColor As Variant) As Integer
    Dim n_sheets, i, j, nTab As Integer
    n_sheets = sheets_count()
    
    nTab = 0
    
    For i = 1 To n_sheets
        If (Sheets(CStr(i)).Tab.color = tabColor) Then
            nTab = nTab + 1
        End If
    Next i
    
    GetLengthByColor = nTab
End Function

Sub get_tabsize_by_well(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Variant, ByRef n_tabcolors As Variant)
    ' n_tabcolors : return value
    ' nof_unique_tab : return value
    
    Dim n_sheets, i, j As Integer
    Dim limit()     As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim limit(0 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    
    For i = 0 To UBound(new_tabcolors)
        limit(i) = GetLengthByColor(new_tabcolors(i))
    Next i
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
    n_tabcolors = new_tabcolors
End Sub
