
Private Sub Workbook_Open()

   
End Sub
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
Private driver As Selenium.chromeDriver


Public Function StringToIntArray(str As String) As Variant
    Dim temp As String, i As Long, L As Long
    Dim CH As String
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction

    temp = ""
    L = Len(str)
    For i = 1 To L
        CH = Mid(str, i, 1)
        If CH Like "[0-9]" Then
            temp = temp & CH
        Else
            temp = temp & " "
        End If
    Next i

    StringToIntArray = Split(wf.Trim(temp), " ")
End Function

Public Function StringToDoubleArray(str As String) As Variant
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    
    Dim trimString As String

    trimString = LTrim(RTrim(str))
   
    StringToDoubleArray = Split(trimString, vbLf)
End Function


'
' minhwasoo, 2024-01-01
' error checking
'
Private Sub delete_ignore_error()
    
    Dim rg1 As Range
    
    For Each rg1 In Range("o6:o35")
            If rg1.Errors.item(xlOmittedCells).Ignore = False Then
                rg1.Errors.item(xlOmittedCells).Ignore = True
            End If
            
            rg1.Errors.item(xlInconsistentFormula).Ignore = True
             
            Debug.Print "a"
    Next rg1

    For Each rg1 In Range("o44:o53")
            If rg1.Errors.item(xlOmittedCells).Ignore = False Then
                rg1.Errors.item(xlOmittedCells).Ignore = True
            End If
            
            rg1.Errors.item(xlInconsistentFormula).Ignore = True
    Next rg1

End Sub

'쉬트를 순회하면서, 에러를 지운다.
Sub deleteall_igonre_error()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Sheets
        'MsgBox ws.name
        ws.Activate
        Call delete_ignore_error
    Next ws

End Sub


Sub ChangeFormat()
    
    Dim lang_code As Integer
    Dim str_format As String

    lang_code = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    

    ' 1042 - korean
    ' 1033 - english
    
    If lang_code = 1042 Then
        str_format = "빨강"
    Else
         str_format = "Red"
    End If

    Range("B6:N35").Select
    Selection.NumberFormatLocal = "0_);[" & str_format & "](0)"
    Selection.NumberFormatLocal = "0.0_);[" & str_format & "](0.0)"

    Range("B6:B35").Select
    Selection.NumberFormatLocal = "0_);[" & str_format & "](0)"
    
End Sub


Sub clear_30year_data()
    Range("b6:n35").ClearContents
End Sub


Function get_area_code() As Integer
    get_area_code = Sheets("main").Range("local_code")
End Function



Sub get_weather_data()
    Dim driver As New chromeDriver
    Dim ddl As Selenium.SelectElement
    
    Dim url As String
    Dim one_string, two_string As String
    Dim sYear, eYear As Integer
    Dim str As String
    
    Range("B2").Value = "30년 " & Range("S8").Value & "데이터, " & Now()
    
    url = "https://data.kma.go.kr/stcs/grnd/grndRnList.do?pgmNo=69"
    Set driver = New Selenium.chromeDriver
    
    driver.SetBinary "c:\ProgramData\00_chrome\chrome.exe" ' Update this path
    
    driver.Start
    driver.AddArgument "--headless"
    driver.Window.Maximize
    driver.Get url

    Sleep (2 * 1000)
        
    '2023/10/28 일, 홈페이지 코드가 변경됨 ...
    ' id="ztree_61_switch"
    ' <a href="javascript:;" id="ztree_61_switch" onclick="treeBtChange(this)" class="button level1 switch center_close" treenode_switch=""><span class="blind">열기</span></a>
    ' #ztree_61_switch, selector복사로 취득
    
    one_string = "ztree_" & CStr(Range("S10").Value) & "_switch"
       
    
    If Range("R9").Value = "Table7" Then
        two_string = Range("S8").Value
    Else
        two_string = Range("S8").Value & " (" & CStr(Range("S9").Value) & ")"
    End If
    
    '금산 (238)
        
    Set ddl = driver.FindElementByCss("#dataFormCd").AsSelect
    ddl.SelectByText ("월")
    Sleep (0.5 * 1000)
    
    
    ' ---------------------------------------------------------------
    
    driver.FindElementByCss("#txtStnNm").Click
    Sleep (1 * 1000)
    
    driver.FindElementByCss("#" & one_string).Click
    Sleep (1 * 1000)
    
    driver.FindElementByLinkText(two_string).Click
    Sleep (1 * 1000)
    
    driver.FindElementByLinkText("선택완료").Click
    
    
    ' ---------------------------------------------------------------
    ' 시작년도, 끝년도 삽입
    
    eYear = Year(Now()) - 1
    sYear = eYear - 29
    
    Set ddl = driver.FindElementByCss("#startYear").AsSelect
    ddl.SelectByText (CStr(sYear))
    Sleep (0.5 * 1000)
   
    Set ddl = driver.FindElementByCss("#endYear").AsSelect
    ddl.SelectByText (CStr(eYear))
    Sleep (0.5 * 1000)
    ' ---------------------------------------------------------------
    
    ' Search Button
    ' driver.FindElementByXPath("//*[@id='schForm']/div[2]").Click
    ' copy by selector
    
    '검색 버튼클릭
    ' driver.FindElementByCss("#schForm > div.wrap_btn > button").Click
    driver.FindElementByCss("button.SEARCH_BTN").Click
    

    Sleep (2 * 1000)
    
    ' Excel download button
    ' driver.FindElementByLinkText("Excel").Click
     
     
    'Excel download
    ' driver.FindElementByCss("#wrap_content > div:nth-child(15) > div.hd_itm > div > a.DOWNLOAD_BTN_XLS").Click
      
    'CSV download
    ' driver.FindElementByCss("#wrap_content > div:nth-child(15) > div.hd_itm > div > a.DOWNLOAD_BTN").Click
    driver.FindElementByCss("a.DOWNLOAD_BTN").Click
    
    
    Sleep (3 * 1000)

End Sub



Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Option Explicit

Sub BackupData()

    Sheets("main").Select
    Sheets("main").Copy Before:=Sheets(1)
    ActiveWindow.SmallScroll Down:=-15
    
    
    ActiveSheet.Shapes.Range(Array("CommandButton_AnnualReset")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton_BackUP")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton_Clear30Year")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton_GetWeatherData")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton_LoadDataFromArray")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton_DeleteIgnoreError")).Select
    Selection.Delete
    
       
    Columns("R:X").Select
    Selection.Delete Shift:=xlToLeft
    Range("Q10").Select
    

    On Error GoTo Catch
    ActiveSheet.name = Range("'main'!$S$8").Value
    Range("B2").Value = Range("'main'!$S$8").Value & " Data, -- " & Now()
    SetRandomSheetTabColor
    
Catch:
    Exit Sub
    
End Sub


Sub SetRandomSheetTabColor()

    ' Define an array of RGB colors
    Dim Colors(1 To 10) As Long
    
    Colors(1) = RGB(192, 0, 0) ' Red
    Colors(2) = RGB(255, 165, 0) ' Orange
    Colors(3) = RGB(255, 255, 0) ' Yellow
    Colors(4) = RGB(0, 176, 80) ' Green
    Colors(5) = RGB(0, 112, 192) ' Blue
    Colors(6) = RGB(112, 48, 160) ' Purple
    Colors(7) = RGB(128, 128, 128) ' Gray
    Colors(8) = RGB(255, 192, 203) ' Pink
    Colors(9) = RGB(128, 64, 64) ' brown
    Colors(10) = RGB(64, 224, 208) ' turquoise

    ' Generate a random number to select a color from the array
    Randomize
    Dim RandomIndex As Integer
    RandomIndex = Int((UBound(Colors) + 1) * Rnd)
    
    ' Set the sheet tab color to the randomly selected color
    ActiveSheet.Tab.Color = Colors(RandomIndex)

End Sub


Sub ShiftUp()

    Range("B7:N35").Select
    Selection.Copy
    
    Range("B6").Select
    ActiveSheet.PasteSpecial Format:=3, link:=1, DisplayAsIcon:=False, IconFileName:=False
    
    Range("B35:N35").Select
    Selection.ClearContents
    
End Sub


Sub CopySingleData()

    Dim i As Integer
    
    For i = 0 To 12
        Cells(35, i + 2).Value = Sheets("main").Cells(40, i + 2).Value
    Next i

End Sub


Sub ShiftNewYear()

    Dim nYear As Integer

    nYear = Year(Now()) - 30
    
    If Range("B6").Value = nYear Then
        Exit Sub
    End If
    
    Call ShiftUp
    Call get_single_data

End Sub



Function get_currentarea_code() As Integer

    Dim name As String
    Dim nArea As Integer
    Dim tbl As ListObject
    
    On Error GoTo Process
    
    Set tbl = Sheets("Code").ListObjects("tblCode")
    name = ActiveSheet.name
   
    nArea = Application.WorksheetFunction.VLookup(name, tbl.Range, 2, False)
    get_currentarea_code = nArea
    Exit Function
    
Process:
    nArea = 0
    get_currentarea_code = nArea
        
End Function



Sub TransPose30Year()
    
    Dim i, j        As Integer
    Dim i1, i2      As Integer
    Dim sYear, eYear As Integer
    
    Range("C1").Select
    Selection.End(xlDown).Select
    
    eYear = Year(Now()) - 1
    sYear = eYear - 29
    
    For i = 1 To 30
        
        i1 = 12 * (i - 1) + 9
        i2 = i1 + 11
        
        Range("C" & CStr(i1) & ":C" & CStr(i2)).Select
        Selection.Copy
        
        j = i + 8
        Range("G" & CStr(j)).Select
        
        Range("F" & CStr(j)).Value = sYear + i - 1
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                               False, Transpose:=True
        
    Next i
    
End Sub


Public Function MyDocPath() As String
    MyDocPath = Environ$("USERPROFILE") & "\" & "Downloads" & "\"
    Debug.Print MyDocsPath
End Function

Public Function SelectDocument() As String
    On Error GoTo Trap
    
    Dim fd          As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .InitialFileName = MyDocPath
        
        .Title = "Please Select the file."
        .Filters.Clear
        .Filters.Add "Excel 2003", "*.xls?"
    End With
    
    'if a selection was made, return the file path
    If fd.Show = -1 Then SelectDocument = fd.SelectedItems(1)
    
Leave:
    Set fd = Nothing
    On Error GoTo 0
    Exit Function
    
Trap:
    MsgBox Err.Description, vbCritical
    Resume Leave
End Function

Function GetFileSize(strFile As Variant) As Long
    Dim lngFSize    As Long, lngDSize As Long
    Dim oFO         As Object
    Dim OFS         As Object
    
    lngFSize = 0
    Set OFS = CreateObject("Scripting.FileSystemObject")
    
    If OFS.FileExists(strFile) Then
        Set oFO = OFS.Getfile(strFile)
        GetFileSize = oFO.Size
    End If
    
End Function

Sub import30YearData()
    
    Dim directory   As String, fileName As String
    Dim fd          As Office.FileDialog
    
    Dim file_name   As String
    Dim thisFileName As String
    file_name = SelectDocument()
    
    Debug.Print getFileName(file_name)
    Debug.Print getPath(file_name)
    
    thisFileName = ActiveWorkbook.name
    
    If file_name <> vbNullString Then
        Workbooks.Open fileName:=file_name
    End If
    
    If file_name = "" Then Exit Sub
    If GetFileSize(file_name) < 5000 Then
        file_name = getFileName(file_name)
        Workbooks(file_name).Close SaveChanges:=False
        Exit Sub
    End If
    
    file_name = getFileName(file_name)
    Call TransPose30Year
    
    Range("F9:R38").Select
    Selection.Copy
    
    Windows(thisFileName).Activate
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Workbooks(file_name).Close SaveChanges:=False
    Application.CutCopyMode = False
    
End Sub


Function OpenRecentFile() As String
    
    Dim strFilePath As String
    Dim fileName As String
    
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim aStack As Object
    
    Dim lngFileCount As Long
    Dim lngCounter As Long

    strFilePath = Environ("USERPROFILE") & "\Downloads\"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(strFilePath)
    Set aStack = CreateObject("System.Collections.Stack")
    
    Debug.Print strFilePath
     
    Set objFolder = objFSO.GetFolder(strFilePath)
    lngFileCount = objFolder.Files.Count
               
    For Each objFile In objFolder.Files
        aStack.Push (objFile.name)
        Debug.Print objFile.name
    Next objFile
        
        
    For lngCounter = 1 To lngFileCount
    
        fileName = aStack.pop
               
        If Right(fileName, 4) = ".xls" Or Right(fileName, 4) = ".csv" Then
            Workbooks.Open fileName:=strFilePath & fileName
            OpenRecentFile = fileName
            Exit For
        End If

    Next lngCounter

    Set objFile = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing
    Set aStack = Nothing
    
End Function


Sub import30RecentData()

    Dim thisFileName, file_name As String
    
    thisFileName = ActiveWorkbook.name
    
    file_name = OpenRecentFile
    Call TransPose30Year
    
    Range("F9:R38").Select
    Selection.Copy
    
    Windows(thisFileName).Activate
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
    Workbooks(file_name).Close SaveChanges:=False
    
    Application.CutCopyMode = False
   
    
End Sub



Function getPath(FullPath As String, Optional Delim As String = "\") As String
    Dim a
    a = Split(FullPath & "$", Delim)
    getPath = Join(Filter(a, a(UBound(a)), False), Delim)
End Function

Function getFileName(FullPath As String) As String
    getFileName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
End Function

Function GetDirOrFileSize(strFolder As String, Optional strFile As Variant) As Long
    
    Dim lngFSize    As Long, lngDSize As Long
    Dim oFO         As Object
    Dim oFD         As Object
    Dim OFS         As Object
    
    lngFSize = 0
    Set OFS = CreateObject("Scripting.FileSystemObject")
    
    If strFolder = "" Then strFolder = ActiveWorkbook.Path
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    'Thanks to Jean-Francois Corbett, you can use also OFS.BuildPath(strFolder, strFile)
    
    If OFS.FolderExists(strFolder) Then
        If Not IsMissing(strFile) Then
            
            If OFS.FileExists(strFolder & strFile) Then
                Set oFO = OFS.Getfile(strFolder & strFile)
                GetDirOrFileSize = oFO.Size
            End If
            
        Else
            Set oFD = OFS.GetFolder(strFolder)
            GetDirOrFileSize = oFD.Size
        End If
        
    End If
    
End Function

Sub CallOpenAPI()
    Dim strURL      As String
    Dim strResult   As String
    
    Dim objHttp     As New WinHttpRequest
    
    strURL = "Open API 주소를 입력하세요"
    objHttp.Open "GET", strURL, False
    objHttp.send
    
    If objHttp.Status = 200 Then        '성공했을 경우
    strResult = objHttp.responseText
    
    'XML로 연결
    Dim objXml      As MSXML2.DOMDocument60
    Set objXml = New DOMDocument60
    objXml.LoadXML (strResult)
    
    '노드 연결
    Dim nodeList    As IXMLDOMNodeList
    Dim nodeRow     As IXMLDOMNode
    Dim nodeCell    As IXMLDOMNode
    Dim nRowCount   As Integer
    Dim nCellCount  As Integer
    
    Set nodeList = objXml.SelectNodes("/response/fields/field")
    
    nRowCount = Range("A60000").End(xlUp).Row
    For Each nodeRow In nodeList
        nRowCount = nRowCount + 1
        
        nCellCount = 0
        For Each nodeCell In nodeRow.ChildNodes
            nCellCount = nCellCount + 1
            '엑셀에 값 반영
            Cells(nRowCount, nCellCount).Value = nodeCell.Text
        Next nodeCell
        
    Next nodeRow
    
Else
    MsgBox "접속에 에러가 발생했습니다"
    
End If
End Sub

Private Sub DumpRangeToArray()
    Dim myArray As Variant
    Dim rng As Range
    Dim cell As Range
    Dim i As Integer, j As Integer

    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.Value
    
    ' Loop through the array (for demonstration purposes)
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Debug.Print myArray(i, j)
        Next j
    Next i
End Sub


Private Sub DumpRangeToArrayAndSaveLoad()
    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.Value
    
    ' Save array to a file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\myArray.csv"
    SaveArrayToFile myArray, filePath
    
    ' Load array from file
    Dim loadedArray() As Variant
    Dim finalArray() As Variant
    
    loadedArray = LoadArrayFromFile(filePath)
    
    'ReDim loadedArray(1 To 30, 1 To 13)
    ' Check if the loaded array is the same as the original array
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            If myArray(i, j) <> CDbl(loadedArray(i, j)) Then
                MsgBox "Arrays are not equal!" & "i :" & i & "  j : " & j
                Exit Sub
            End If
        Next j
    Next i
    
    MsgBox "Arrays are equal!"
End Sub


Private Sub SaveArrayToFile(myArray As Variant, filePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Print #FileNum, myArray(i, j);
            
            ' Separate values with a comma (CSV format)
            If j < UBound(myArray, 2) Then
                Print #FileNum, ",";
            End If
        Next j
        ' Start a new line for each row
        Print #FileNum, ""
    Next i
    
    Close FileNum
End Sub

Private Function LoadArrayFromFile(filePath As String) As Variant
    Dim FileContent As String
    Dim Lines() As String
    Dim Values() As String
    Dim i As Integer, j As Integer
    Dim loadedArray() As Variant
    
    Open filePath For Input As #1
    FileContent = Input$(LOF(1), #1)
    Close #1
    
    Lines = Split(FileContent, vbCrLf)
    ReDim loadedArray(1 To UBound(Lines) + 1, 1 To UBound(Split(Lines(0), ",")) + 1)
    
    For i = LBound(Lines) To UBound(Lines)
        Values = Split(Lines(i), ",")
        For j = LBound(Values) To UBound(Values)
            loadedArray(i + 1, j + 1) = Values(j)
        Next j
    Next i
    
    LoadArrayFromFile = loadedArray
End Function
' year 2023
'data_GEUMSAN
'data_BORYUNG
'data_DAEJEON
'data_BUYEO
'data_SEOSAN
'data_CHEONAN
'data_CHEUNGJU

Function data_TEMP() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    data_TEMP = myArray

End Function

Function data_CHUNGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 11.9
myArray(1, 3) = 6.8
myArray(1, 4) = 29.3
myArray(1, 5) = 46.2
myArray(1, 6) = 45.5
myArray(1, 7) = 19
myArray(1, 8) = 290.8
myArray(1, 9) = 802
myArray(1, 10) = 16
myArray(1, 11) = 23
myArray(1, 12) = 23.9
myArray(1, 13) = 4.2

myArray(2, 1) = 1996
myArray(2, 2) = 27.8
myArray(2, 3) = 1.6
myArray(2, 4) = 101.7
myArray(2, 5) = 36.5
myArray(2, 6) = 29.1
myArray(2, 7) = 203.5
myArray(2, 8) = 207
myArray(2, 9) = 126
myArray(2, 10) = 26.5
myArray(2, 11) = 83
myArray(2, 12) = 74.1
myArray(2, 13) = 22.2

myArray(3, 1) = 1997
myArray(3, 2) = 5
myArray(3, 3) = 44.3
myArray(3, 4) = 23.2
myArray(3, 5) = 60
myArray(3, 6) = 193.5
myArray(3, 7) = 147.8
myArray(3, 8) = 308
myArray(3, 9) = 179.9
myArray(3, 10) = 52.5
myArray(3, 11) = 35.6
myArray(3, 12) = 128.5
myArray(3, 13) = 41.4

myArray(4, 1) = 1998
myArray(4, 2) = 22.5
myArray(4, 3) = 26.4
myArray(4, 4) = 33.8
myArray(4, 5) = 145.3
myArray(4, 6) = 90
myArray(4, 7) = 217.7
myArray(4, 8) = 286.6
myArray(4, 9) = 541.7
myArray(4, 10) = 183
myArray(4, 11) = 64.5
myArray(4, 12) = 38.5
myArray(4, 13) = 2.5

myArray(5, 1) = 1999
myArray(5, 2) = 3.1
myArray(5, 3) = 1.9
myArray(5, 4) = 54.2
myArray(5, 5) = 96
myArray(5, 6) = 103.2
myArray(5, 7) = 168.5
myArray(5, 8) = 112.7
myArray(5, 9) = 299.6
myArray(5, 10) = 239.6
myArray(5, 11) = 195.1
myArray(5, 12) = 34.5
myArray(5, 13) = 7

myArray(6, 1) = 2000
myArray(6, 2) = 47.5
myArray(6, 3) = 3.3
myArray(6, 4) = 14.5
myArray(6, 5) = 42.5
myArray(6, 6) = 54
myArray(6, 7) = 248.5
myArray(6, 8) = 259.5
myArray(6, 9) = 260.5
myArray(6, 10) = 260.8
myArray(6, 11) = 25.5
myArray(6, 12) = 35.5
myArray(6, 13) = 17.5

myArray(7, 1) = 2001
myArray(7, 2) = 47
myArray(7, 3) = 52.2
myArray(7, 4) = 8.3
myArray(7, 5) = 12
myArray(7, 6) = 4.6
myArray(7, 7) = 238.9
myArray(7, 8) = 241.6
myArray(7, 9) = 82.6
myArray(7, 10) = 13.8
myArray(7, 11) = 82.1
myArray(7, 12) = 4.4
myArray(7, 13) = 10.6

myArray(8, 1) = 2002
myArray(8, 2) = 49.6
myArray(8, 3) = 4.2
myArray(8, 4) = 28.5
myArray(8, 5) = 151.2
myArray(8, 6) = 105
myArray(8, 7) = 74.7
myArray(8, 8) = 190.2
myArray(8, 9) = 653
myArray(8, 10) = 92.6
myArray(8, 11) = 52.2
myArray(8, 12) = 9.8
myArray(8, 13) = 58.6

myArray(9, 1) = 2003
myArray(9, 2) = 17.2
myArray(9, 3) = 59.2
myArray(9, 4) = 58.2
myArray(9, 5) = 170.8
myArray(9, 6) = 117.8
myArray(9, 7) = 152.1
myArray(9, 8) = 382.8
myArray(9, 9) = 314.7
myArray(9, 10) = 268.1
myArray(9, 11) = 27.5
myArray(9, 12) = 55
myArray(9, 13) = 17.8

myArray(10, 1) = 2004
myArray(10, 2) = 16.6
myArray(10, 3) = 32
myArray(10, 4) = 29.4
myArray(10, 5) = 81
myArray(10, 6) = 124.9
myArray(10, 7) = 335
myArray(10, 8) = 410.7
myArray(10, 9) = 192.2
myArray(10, 10) = 144.1
myArray(10, 11) = 1.4
myArray(10, 12) = 32.5
myArray(10, 13) = 25.4

myArray(11, 1) = 2005
myArray(11, 2) = 2.8
myArray(11, 3) = 20.8
myArray(11, 4) = 43.1
myArray(11, 5) = 63.1
myArray(11, 6) = 53.9
myArray(11, 7) = 178.7
myArray(11, 8) = 381.6
myArray(11, 9) = 226.1
myArray(11, 10) = 320
myArray(11, 11) = 63.4
myArray(11, 12) = 15.7
myArray(11, 13) = 11.7

myArray(12, 1) = 2006
myArray(12, 2) = 27.1
myArray(12, 3) = 34.9
myArray(12, 4) = 5.9
myArray(12, 5) = 91.8
myArray(12, 6) = 95.1
myArray(12, 7) = 128.5
myArray(12, 8) = 666.9
myArray(12, 9) = 71.5
myArray(12, 10) = 21.7
myArray(12, 11) = 23.1
myArray(12, 12) = 53.1
myArray(12, 13) = 14.3

myArray(13, 1) = 2007
myArray(13, 2) = 5.5
myArray(13, 3) = 38.5
myArray(13, 4) = 112.7
myArray(13, 5) = 18.3
myArray(13, 6) = 116.5
myArray(13, 7) = 90.1
myArray(13, 8) = 282.7
myArray(13, 9) = 366
myArray(13, 10) = 332.7
myArray(13, 11) = 32.8
myArray(13, 12) = 22
myArray(13, 13) = 21.4

myArray(14, 1) = 2008
myArray(14, 2) = 29.3
myArray(14, 3) = 8.2
myArray(14, 4) = 43.1
myArray(14, 5) = 31.5
myArray(14, 6) = 70.9
myArray(14, 7) = 78.1
myArray(14, 8) = 319.8
myArray(14, 9) = 192.5
myArray(14, 10) = 71.1
myArray(14, 11) = 16
myArray(14, 12) = 10.3
myArray(14, 13) = 11.7

myArray(15, 1) = 2009
myArray(15, 2) = 16.7
myArray(15, 3) = 15.8
myArray(15, 4) = 52
myArray(15, 5) = 30.7
myArray(15, 6) = 97.1
myArray(15, 7) = 89.5
myArray(15, 8) = 316.2
myArray(15, 9) = 142.5
myArray(15, 10) = 70.6
myArray(15, 11) = 45
myArray(15, 12) = 31.2
myArray(15, 13) = 29.5

myArray(16, 1) = 2010
myArray(16, 2) = 44.3
myArray(16, 3) = 70.8
myArray(16, 4) = 85.3
myArray(16, 5) = 69.5
myArray(16, 6) = 97
myArray(16, 7) = 50.6
myArray(16, 8) = 112.2
myArray(16, 9) = 345.1
myArray(16, 10) = 287.8
myArray(16, 11) = 21
myArray(16, 12) = 14.3
myArray(16, 13) = 14.4

myArray(17, 1) = 2011
myArray(17, 2) = 2.7
myArray(17, 3) = 45.9
myArray(17, 4) = 30.6
myArray(17, 5) = 157.8
myArray(17, 6) = 187.7
myArray(17, 7) = 452.6
myArray(17, 8) = 603.9
myArray(17, 9) = 289.4
myArray(17, 10) = 158.6
myArray(17, 11) = 61.5
myArray(17, 12) = 67
myArray(17, 13) = 15.6

myArray(18, 1) = 2012
myArray(18, 2) = 9.6
myArray(18, 3) = 1.7
myArray(18, 4) = 66.4
myArray(18, 5) = 84.5
myArray(18, 6) = 61
myArray(18, 7) = 58.8
myArray(18, 8) = 265.7
myArray(18, 9) = 403.3
myArray(18, 10) = 177.2
myArray(18, 11) = 62
myArray(18, 12) = 48
myArray(18, 13) = 52.1

myArray(19, 1) = 2013
myArray(19, 2) = 40.5
myArray(19, 3) = 36.9
myArray(19, 4) = 48
myArray(19, 5) = 84.7
myArray(19, 6) = 92.5
myArray(19, 7) = 126.6
myArray(19, 8) = 240.7
myArray(19, 9) = 222.2
myArray(19, 10) = 122.2
myArray(19, 11) = 12.1
myArray(19, 12) = 44
myArray(19, 13) = 32.2

myArray(20, 1) = 2014
myArray(20, 2) = 14
myArray(20, 3) = 18.9
myArray(20, 4) = 37.7
myArray(20, 5) = 39.6
myArray(20, 6) = 26.3
myArray(20, 7) = 63.3
myArray(20, 8) = 92.6
myArray(20, 9) = 284.3
myArray(20, 10) = 122.7
myArray(20, 11) = 153.8
myArray(20, 12) = 23.5
myArray(20, 13) = 22.9

myArray(21, 1) = 2015
myArray(21, 2) = 15.6
myArray(21, 3) = 22.8
myArray(21, 4) = 31.7
myArray(21, 5) = 88.9
myArray(21, 6) = 23
myArray(21, 7) = 75
myArray(21, 8) = 181.6
myArray(21, 9) = 71.8
myArray(21, 10) = 33.8
myArray(21, 11) = 60.2
myArray(21, 12) = 89.9
myArray(21, 13) = 37.5

myArray(22, 1) = 2016
myArray(22, 2) = 1.8
myArray(22, 3) = 50.1
myArray(22, 4) = 11.9
myArray(22, 5) = 97.3
myArray(22, 6) = 70
myArray(22, 7) = 38.9
myArray(22, 8) = 374.4
myArray(22, 9) = 44
myArray(22, 10) = 60.8
myArray(22, 11) = 102.6
myArray(22, 12) = 22.9
myArray(22, 13) = 42.4

myArray(23, 1) = 2017
myArray(23, 2) = 18
myArray(23, 3) = 36.2
myArray(23, 4) = 22.5
myArray(23, 5) = 71.4
myArray(23, 6) = 32.2
myArray(23, 7) = 43.7
myArray(23, 8) = 464.3
myArray(23, 9) = 257.9
myArray(23, 10) = 62.4
myArray(23, 11) = 21.2
myArray(23, 12) = 17.6
myArray(23, 13) = 25.5

myArray(24, 1) = 2018
myArray(24, 2) = 14.3
myArray(24, 3) = 35.8
myArray(24, 4) = 75.1
myArray(24, 5) = 107.9
myArray(24, 6) = 180
myArray(24, 7) = 63.7
myArray(24, 8) = 149.1
myArray(24, 9) = 353.3
myArray(24, 10) = 184.9
myArray(24, 11) = 96
myArray(24, 12) = 50
myArray(24, 13) = 39

myArray(25, 1) = 2019
myArray(25, 2) = 4.1
myArray(25, 3) = 29
myArray(25, 4) = 27.9
myArray(25, 5) = 58.5
myArray(25, 6) = 15.4
myArray(25, 7) = 59.6
myArray(25, 8) = 161.4
myArray(25, 9) = 102.6
myArray(25, 10) = 165.9
myArray(25, 11) = 59
myArray(25, 12) = 84.6
myArray(25, 13) = 27.5

myArray(26, 1) = 2020
myArray(26, 2) = 60.2
myArray(26, 3) = 61.9
myArray(26, 4) = 20.7
myArray(26, 5) = 25.9
myArray(26, 6) = 109.7
myArray(26, 7) = 112.1
myArray(26, 8) = 352.2
myArray(26, 9) = 505.6
myArray(26, 10) = 146.4
myArray(26, 11) = 10.7
myArray(26, 12) = 30
myArray(26, 13) = 11.1

myArray(27, 1) = 2021
myArray(27, 2) = 13.6
myArray(27, 3) = 12.3
myArray(27, 4) = 80.5
myArray(27, 5) = 63.4
myArray(27, 6) = 178.4
myArray(27, 7) = 130.4
myArray(27, 8) = 310.7
myArray(27, 9) = 239.9
myArray(27, 10) = 240.3
myArray(27, 11) = 45.5
myArray(27, 12) = 44.9
myArray(27, 13) = 5.6

myArray(28, 1) = 2022
myArray(28, 2) = 2
myArray(28, 3) = 4.3
myArray(28, 4) = 79.9
myArray(28, 5) = 45.8
myArray(28, 6) = 8.6
myArray(28, 7) = 219
myArray(28, 8) = 350.5
myArray(28, 9) = 457.3
myArray(28, 10) = 102
myArray(28, 11) = 96.5
myArray(28, 12) = 79.5
myArray(28, 13) = 18.8

myArray(29, 1) = 2023
myArray(29, 2) = 23
myArray(29, 3) = 3.3
myArray(29, 4) = 16.7
myArray(29, 5) = 32.1
myArray(29, 6) = 123.8
myArray(29, 7) = 239.6
myArray(29, 8) = 554.2
myArray(29, 9) = 248
myArray(29, 10) = 242.5
myArray(29, 11) = 31.3
myArray(29, 12) = 50.8
myArray(29, 13) = 96.3

myArray(30, 1) = 2024
myArray(30, 2) = 34.6
myArray(30, 3) = 86.6
myArray(30, 4) = 50.8
myArray(30, 5) = 59
myArray(30, 6) = 128.7
myArray(30, 7) = 126.1
myArray(30, 8) = 503.2
myArray(30, 9) = 71.4
myArray(30, 10) = 211.4
myArray(30, 11) = 122.8
myArray(30, 12) = 49.1
myArray(30, 13) = 4.3


    data_CHUNGJU = myArray

End Function


Function data_CHUPUNGNYEONG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 16.7
myArray(1, 3) = 15.7
myArray(1, 4) = 55.6
myArray(1, 5) = 67.5
myArray(1, 6) = 37
myArray(1, 7) = 43.8
myArray(1, 8) = 132.2
myArray(1, 9) = 472.4
myArray(1, 10) = 56.9
myArray(1, 11) = 22
myArray(1, 12) = 29.1
myArray(1, 13) = 5.1

myArray(2, 1) = 1996
myArray(2, 2) = 19.4
myArray(2, 3) = 2
myArray(2, 4) = 117.2
myArray(2, 5) = 28.1
myArray(2, 6) = 58.9
myArray(2, 7) = 463.9
myArray(2, 8) = 118.2
myArray(2, 9) = 89.4
myArray(2, 10) = 18.1
myArray(2, 11) = 73.6
myArray(2, 12) = 58.9
myArray(2, 13) = 24.3

myArray(3, 1) = 1997
myArray(3, 2) = 15.2
myArray(3, 3) = 32.8
myArray(3, 4) = 27.9
myArray(3, 5) = 49.7
myArray(3, 6) = 148.7
myArray(3, 7) = 205.2
myArray(3, 8) = 249.2
myArray(3, 9) = 154.4
myArray(3, 10) = 30
myArray(3, 11) = 4.2
myArray(3, 12) = 140.7
myArray(3, 13) = 44.6

myArray(4, 1) = 1998
myArray(4, 2) = 28.2
myArray(4, 3) = 49.6
myArray(4, 4) = 30.9
myArray(4, 5) = 201.5
myArray(4, 6) = 87.8
myArray(4, 7) = 227
myArray(4, 8) = 244.8
myArray(4, 9) = 368.9
myArray(4, 10) = 282.7
myArray(4, 11) = 49.9
myArray(4, 12) = 15.2
myArray(4, 13) = 4.4

myArray(5, 1) = 1999
myArray(5, 2) = 2.8
myArray(5, 3) = 20.1
myArray(5, 4) = 86.7
myArray(5, 5) = 76.4
myArray(5, 6) = 113.2
myArray(5, 7) = 167.3
myArray(5, 8) = 143.3
myArray(5, 9) = 240.1
myArray(5, 10) = 294.5
myArray(5, 11) = 93.5
myArray(5, 12) = 16.3
myArray(5, 13) = 15.3

myArray(6, 1) = 2000
myArray(6, 2) = 29.3
myArray(6, 3) = 5
myArray(6, 4) = 26
myArray(6, 5) = 43.4
myArray(6, 6) = 26.1
myArray(6, 7) = 155.7
myArray(6, 8) = 293.4
myArray(6, 9) = 318.5
myArray(6, 10) = 304.6
myArray(6, 11) = 30.4
myArray(6, 12) = 51
myArray(6, 13) = 11.1

myArray(7, 1) = 2001
myArray(7, 2) = 47.7
myArray(7, 3) = 63.9
myArray(7, 4) = 9.7
myArray(7, 5) = 16.5
myArray(7, 6) = 39.6
myArray(7, 7) = 202.6
myArray(7, 8) = 154.2
myArray(7, 9) = 23.4
myArray(7, 10) = 112.9
myArray(7, 11) = 116.5
myArray(7, 12) = 10.7
myArray(7, 13) = 24.5

myArray(8, 1) = 2002
myArray(8, 2) = 71.9
myArray(8, 3) = 8.9
myArray(8, 4) = 54.4
myArray(8, 5) = 188.2
myArray(8, 6) = 125.3
myArray(8, 7) = 49.3
myArray(8, 8) = 204.3
myArray(8, 9) = 597.1
myArray(8, 10) = 55.5
myArray(8, 11) = 39.6
myArray(8, 12) = 19.1
myArray(8, 13) = 46.1

myArray(9, 1) = 2003
myArray(9, 2) = 20.9
myArray(9, 3) = 54.8
myArray(9, 4) = 50.6
myArray(9, 5) = 182.7
myArray(9, 6) = 178.9
myArray(9, 7) = 157.3
myArray(9, 8) = 548.3
myArray(9, 9) = 330.6
myArray(9, 10) = 222.9
myArray(9, 11) = 24
myArray(9, 12) = 47.2
myArray(9, 13) = 17.1

myArray(10, 1) = 2004
myArray(10, 2) = 16
myArray(10, 3) = 24.9
myArray(10, 4) = 20.6
myArray(10, 5) = 70.8
myArray(10, 6) = 112.5
myArray(10, 7) = 249.1
myArray(10, 8) = 391.5
myArray(10, 9) = 317.6
myArray(10, 10) = 175.1
myArray(10, 11) = 2.6
myArray(10, 12) = 40.2
myArray(10, 13) = 23.3

myArray(11, 1) = 2005
myArray(11, 2) = 12.8
myArray(11, 3) = 35.9
myArray(11, 4) = 50.2
myArray(11, 5) = 31.3
myArray(11, 6) = 47
myArray(11, 7) = 131.2
myArray(11, 8) = 252.3
myArray(11, 9) = 291.8
myArray(11, 10) = 107.7
myArray(11, 11) = 13
myArray(11, 12) = 20.6
myArray(11, 13) = 14.6

myArray(12, 1) = 2006
myArray(12, 2) = 19.4
myArray(12, 3) = 30.4
myArray(12, 4) = 8.7
myArray(12, 5) = 89.5
myArray(12, 6) = 102.5
myArray(12, 7) = 128
myArray(12, 8) = 697.6
myArray(12, 9) = 43
myArray(12, 10) = 36.9
myArray(12, 11) = 36
myArray(12, 12) = 61.4
myArray(12, 13) = 19.7

myArray(13, 1) = 2007
myArray(13, 2) = 9.8
myArray(13, 3) = 45.9
myArray(13, 4) = 85
myArray(13, 5) = 24.2
myArray(13, 6) = 73
myArray(13, 7) = 152.1
myArray(13, 8) = 209.5
myArray(13, 9) = 267.9
myArray(13, 10) = 386.1
myArray(13, 11) = 20.4
myArray(13, 12) = 7.2
myArray(13, 13) = 29.9

myArray(14, 1) = 2008
myArray(14, 2) = 40.5
myArray(14, 3) = 6.1
myArray(14, 4) = 27.7
myArray(14, 5) = 47.2
myArray(14, 6) = 57.6
myArray(14, 7) = 172.6
myArray(14, 8) = 152.2
myArray(14, 9) = 172.9
myArray(14, 10) = 59
myArray(14, 11) = 56.5
myArray(14, 12) = 17
myArray(14, 13) = 9.2

myArray(15, 1) = 2009
myArray(15, 2) = 12
myArray(15, 3) = 34.7
myArray(15, 4) = 37.2
myArray(15, 5) = 34
myArray(15, 6) = 112.1
myArray(15, 7) = 87.4
myArray(15, 8) = 436.9
myArray(15, 9) = 97.7
myArray(15, 10) = 54.7
myArray(15, 11) = 16.5
myArray(15, 12) = 52.7
myArray(15, 13) = 34.8

myArray(16, 1) = 2010
myArray(16, 2) = 22.5
myArray(16, 3) = 72.6
myArray(16, 4) = 88.6
myArray(16, 5) = 54.6
myArray(16, 6) = 115.3
myArray(16, 7) = 38.2
myArray(16, 8) = 201.7
myArray(16, 9) = 443.2
myArray(16, 10) = 150
myArray(16, 11) = 22.3
myArray(16, 12) = 15.1
myArray(16, 13) = 36.3

myArray(17, 1) = 2011
myArray(17, 2) = 3.6
myArray(17, 3) = 47.6
myArray(17, 4) = 20.5
myArray(17, 5) = 95.5
myArray(17, 6) = 163.7
myArray(17, 7) = 187.5
myArray(17, 8) = 284.7
myArray(17, 9) = 369.7
myArray(17, 10) = 59.9
myArray(17, 11) = 58.5
myArray(17, 12) = 95.8
myArray(17, 13) = 14.8

myArray(18, 1) = 2012
myArray(18, 2) = 19.8
myArray(18, 3) = 1.6
myArray(18, 4) = 93.8
myArray(18, 5) = 90.6
myArray(18, 6) = 31.3
myArray(18, 7) = 73.8
myArray(18, 8) = 228.6
myArray(18, 9) = 490.1
myArray(18, 10) = 282.6
myArray(18, 11) = 47
myArray(18, 12) = 51
myArray(18, 13) = 55.6

myArray(19, 1) = 2013
myArray(19, 2) = 42.7
myArray(19, 3) = 38.4
myArray(19, 4) = 48.6
myArray(19, 5) = 79.2
myArray(19, 6) = 80.8
myArray(19, 7) = 119.7
myArray(19, 8) = 186.6
myArray(19, 9) = 106.1
myArray(19, 10) = 94.8
myArray(19, 11) = 48
myArray(19, 12) = 48.1
myArray(19, 13) = 27.6

myArray(20, 1) = 2014
myArray(20, 2) = 6.7
myArray(20, 3) = 13.8
myArray(20, 4) = 93.9
myArray(20, 5) = 100.6
myArray(20, 6) = 20.8
myArray(20, 7) = 102.9
myArray(20, 8) = 74.8
myArray(20, 9) = 402.3
myArray(20, 10) = 88.8
myArray(20, 11) = 121.1
myArray(20, 12) = 78.7
myArray(20, 13) = 32

myArray(21, 1) = 2015
myArray(21, 2) = 34.3
myArray(21, 3) = 30.4
myArray(21, 4) = 47.7
myArray(21, 5) = 100.7
myArray(21, 6) = 23.7
myArray(21, 7) = 83.8
myArray(21, 8) = 148.3
myArray(21, 9) = 89.4
myArray(21, 10) = 24.2
myArray(21, 11) = 79.1
myArray(21, 12) = 127.6
myArray(21, 13) = 39.8

myArray(22, 1) = 2016
myArray(22, 2) = 18.9
myArray(22, 3) = 34.2
myArray(22, 4) = 57.5
myArray(22, 5) = 155.8
myArray(22, 6) = 60.1
myArray(22, 7) = 45.5
myArray(22, 8) = 304.4
myArray(22, 9) = 90.6
myArray(22, 10) = 187.2
myArray(22, 11) = 123.8
myArray(22, 12) = 30.2
myArray(22, 13) = 38.9

myArray(23, 1) = 2017
myArray(23, 2) = 20.4
myArray(23, 3) = 45
myArray(23, 4) = 30.9
myArray(23, 5) = 64.7
myArray(23, 6) = 21.3
myArray(23, 7) = 60.6
myArray(23, 8) = 273.4
myArray(23, 9) = 208.2
myArray(23, 10) = 105.9
myArray(23, 11) = 48.2
myArray(23, 12) = 17.2
myArray(23, 13) = 26

myArray(24, 1) = 2018
myArray(24, 2) = 29.2
myArray(24, 3) = 44.3
myArray(24, 4) = 127.7
myArray(24, 5) = 132
myArray(24, 6) = 81.1
myArray(24, 7) = 89.6
myArray(24, 8) = 123.8
myArray(24, 9) = 335.4
myArray(24, 10) = 92.3
myArray(24, 11) = 196.4
myArray(24, 12) = 24.4
myArray(24, 13) = 28

myArray(25, 1) = 2019
myArray(25, 2) = 7.7
myArray(25, 3) = 31.1
myArray(25, 4) = 33.5
myArray(25, 5) = 91.7
myArray(25, 6) = 46.3
myArray(25, 7) = 116.5
myArray(25, 8) = 185.4
myArray(25, 9) = 218.1
myArray(25, 10) = 191.8
myArray(25, 11) = 179.4
myArray(25, 12) = 35.4
myArray(25, 13) = 27.2

myArray(26, 1) = 2020
myArray(26, 2) = 71
myArray(26, 3) = 68.2
myArray(26, 4) = 21
myArray(26, 5) = 39.7
myArray(26, 6) = 65.3
myArray(26, 7) = 182.8
myArray(26, 8) = 499.3
myArray(26, 9) = 333.5
myArray(26, 10) = 226.8
myArray(26, 11) = 5
myArray(26, 12) = 38.3
myArray(26, 13) = 3.1

myArray(27, 1) = 2021
myArray(27, 2) = 18.2
myArray(27, 3) = 14.8
myArray(27, 4) = 83.8
myArray(27, 5) = 43.8
myArray(27, 6) = 165.7
myArray(27, 7) = 50.8
myArray(27, 8) = 217.1
myArray(27, 9) = 285.6
myArray(27, 10) = 161.2
myArray(27, 11) = 41
myArray(27, 12) = 43.9
myArray(27, 13) = 3.2

myArray(28, 1) = 2022
myArray(28, 2) = 3.4
myArray(28, 3) = 2.3
myArray(28, 4) = 71.2
myArray(28, 5) = 57
myArray(28, 6) = 4.7
myArray(28, 7) = 98.5
myArray(28, 8) = 189.1
myArray(28, 9) = 272.6
myArray(28, 10) = 117.1
myArray(28, 11) = 62.5
myArray(28, 12) = 38.7
myArray(28, 13) = 11.5

myArray(29, 1) = 2023
myArray(29, 2) = 23.4
myArray(29, 3) = 8.2
myArray(29, 4) = 32.3
myArray(29, 5) = 43.9
myArray(29, 6) = 144.8
myArray(29, 7) = 254.9
myArray(29, 8) = 390.3
myArray(29, 9) = 288
myArray(29, 10) = 199.8
myArray(29, 11) = 12.2
myArray(29, 12) = 36.6
myArray(29, 13) = 117.1

myArray(30, 1) = 2024
myArray(30, 2) = 42.5
myArray(30, 3) = 100
myArray(30, 4) = 65.6
myArray(30, 5) = 49.3
myArray(30, 6) = 78.5
myArray(30, 7) = 95.4
myArray(30, 8) = 443.6
myArray(30, 9) = 20.6
myArray(30, 10) = 221.3
myArray(30, 11) = 74.9
myArray(30, 12) = 31.2
myArray(30, 13) = 8.6


    data_CHUPUNGNYEONG = myArray

End Function

Function data_CHEONGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 21.5
myArray(1, 3) = 14
myArray(1, 4) = 34.4
myArray(1, 5) = 64
myArray(1, 6) = 70.7
myArray(1, 7) = 30.9
myArray(1, 8) = 204.9
myArray(1, 9) = 835.4
myArray(1, 10) = 17.5
myArray(1, 11) = 22.6
myArray(1, 12) = 20.3
myArray(1, 13) = 3.6

myArray(2, 1) = 1996
myArray(2, 2) = 27.9
myArray(2, 3) = 4.2
myArray(2, 4) = 98.4
myArray(2, 5) = 28.6
myArray(2, 6) = 36.8
myArray(2, 7) = 255.8
myArray(2, 8) = 170.5
myArray(2, 9) = 128.6
myArray(2, 10) = 11.2
myArray(2, 11) = 67.1
myArray(2, 12) = 77.2
myArray(2, 13) = 22.5

myArray(3, 1) = 1997
myArray(3, 2) = 12.9
myArray(3, 3) = 39.1
myArray(3, 4) = 31.6
myArray(3, 5) = 58.5
myArray(3, 6) = 179.1
myArray(3, 7) = 210.3
myArray(3, 8) = 425.5
myArray(3, 9) = 211.1
myArray(3, 10) = 55.5
myArray(3, 11) = 8.4
myArray(3, 12) = 180.3
myArray(3, 13) = 44.3

myArray(4, 1) = 1998
myArray(4, 2) = 22
myArray(4, 3) = 28.9
myArray(4, 4) = 30.9
myArray(4, 5) = 153.1
myArray(4, 6) = 92.8
myArray(4, 7) = 247
myArray(4, 8) = 253
myArray(4, 9) = 460.6
myArray(4, 10) = 225.9
myArray(4, 11) = 74.2
myArray(4, 12) = 44.7
myArray(4, 13) = 7.1

myArray(5, 1) = 1999
myArray(5, 2) = 1.6
myArray(5, 3) = 3.6
myArray(5, 4) = 54.1
myArray(5, 5) = 91.4
myArray(5, 6) = 102.4
myArray(5, 7) = 191.1
myArray(5, 8) = 122.4
myArray(5, 9) = 197.4
myArray(5, 10) = 281.3
myArray(5, 11) = 252.4
myArray(5, 12) = 15.4
myArray(5, 13) = 13.4

myArray(6, 1) = 2000
myArray(6, 2) = 38.7
myArray(6, 3) = 1.3
myArray(6, 4) = 10.4
myArray(6, 5) = 56.1
myArray(6, 6) = 42.1
myArray(6, 7) = 185.7
myArray(6, 8) = 300
myArray(6, 9) = 390.4
myArray(6, 10) = 244.6
myArray(6, 11) = 32.1
myArray(6, 12) = 37.3
myArray(6, 13) = 18.9

myArray(7, 1) = 2001
myArray(7, 2) = 56.9
myArray(7, 3) = 50.3
myArray(7, 4) = 11.3
myArray(7, 5) = 12.7
myArray(7, 6) = 14.3
myArray(7, 7) = 217.5
myArray(7, 8) = 171.5
myArray(7, 9) = 135.5
myArray(7, 10) = 11.8
myArray(7, 11) = 75.9
myArray(7, 12) = 6.9
myArray(7, 13) = 19.5

myArray(8, 1) = 2002
myArray(8, 2) = 58.7
myArray(8, 3) = 9
myArray(8, 4) = 25.9
myArray(8, 5) = 132
myArray(8, 6) = 106.9
myArray(8, 7) = 57.9
myArray(8, 8) = 186.2
myArray(8, 9) = 482.4
myArray(8, 10) = 90.5
myArray(8, 11) = 58
myArray(8, 12) = 26.3
myArray(8, 13) = 48

myArray(9, 1) = 2003
myArray(9, 2) = 16.2
myArray(9, 3) = 45
myArray(9, 4) = 38.9
myArray(9, 5) = 192.7
myArray(9, 6) = 113.5
myArray(9, 7) = 186
myArray(9, 8) = 467.2
myArray(9, 9) = 293.9
myArray(9, 10) = 150.6
myArray(9, 11) = 32.5
myArray(9, 12) = 33.1
myArray(9, 13) = 12.2

myArray(10, 1) = 2004
myArray(10, 2) = 12.5
myArray(10, 3) = 42.3
myArray(10, 4) = 67.3
myArray(10, 5) = 61
myArray(10, 6) = 121.8
myArray(10, 7) = 421.5
myArray(10, 8) = 318.9
myArray(10, 9) = 247.6
myArray(10, 10) = 139
myArray(10, 11) = 2
myArray(10, 12) = 34
myArray(10, 13) = 38

myArray(11, 1) = 2005
myArray(11, 2) = 4.6
myArray(11, 3) = 13.8
myArray(11, 4) = 36.8
myArray(11, 5) = 66.1
myArray(11, 6) = 50.7
myArray(11, 7) = 170
myArray(11, 8) = 373.1
myArray(11, 9) = 334.7
myArray(11, 10) = 295.5
myArray(11, 11) = 54.6
myArray(11, 12) = 16
myArray(11, 13) = 11.3

myArray(12, 1) = 2006
myArray(12, 2) = 20
myArray(12, 3) = 28.9
myArray(12, 4) = 8.2
myArray(12, 5) = 89.3
myArray(12, 6) = 119.4
myArray(12, 7) = 115.5
myArray(12, 8) = 508
myArray(12, 9) = 52
myArray(12, 10) = 18.4
myArray(12, 11) = 21.3
myArray(12, 12) = 83.4
myArray(12, 13) = 16.7

myArray(13, 1) = 2007
myArray(13, 2) = 11.2
myArray(13, 3) = 33.3
myArray(13, 4) = 103.2
myArray(13, 5) = 35.8
myArray(13, 6) = 145.5
myArray(13, 7) = 81.2
myArray(13, 8) = 273.2
myArray(13, 9) = 385.5
myArray(13, 10) = 391.4
myArray(13, 11) = 43.5
myArray(13, 12) = 8.8
myArray(13, 13) = 21.9

myArray(14, 1) = 2008
myArray(14, 2) = 29
myArray(14, 3) = 7.7
myArray(14, 4) = 29.4
myArray(14, 5) = 27
myArray(14, 6) = 64.5
myArray(14, 7) = 112
myArray(14, 8) = 296.6
myArray(14, 9) = 195.5
myArray(14, 10) = 92.6
myArray(14, 11) = 13.1
myArray(14, 12) = 10.5
myArray(14, 13) = 14.4

myArray(15, 1) = 2009
myArray(15, 2) = 17.8
myArray(15, 3) = 13.1
myArray(15, 4) = 54.9
myArray(15, 5) = 30.4
myArray(15, 6) = 109.6
myArray(15, 7) = 77.2
myArray(15, 8) = 345.7
myArray(15, 9) = 187.5
myArray(15, 10) = 49.5
myArray(15, 11) = 49.5
myArray(15, 12) = 43.9
myArray(15, 13) = 40.7

myArray(16, 1) = 2010
myArray(16, 2) = 37.8
myArray(16, 3) = 69.2
myArray(16, 4) = 99.8
myArray(16, 5) = 70.5
myArray(16, 6) = 110
myArray(16, 7) = 42.6
myArray(16, 8) = 224.1
myArray(16, 9) = 433.2
myArray(16, 10) = 278.6
myArray(16, 11) = 17.1
myArray(16, 12) = 15.7
myArray(16, 13) = 23.8

myArray(17, 1) = 2011
myArray(17, 2) = 4.5
myArray(17, 3) = 43.2
myArray(17, 4) = 23.5
myArray(17, 5) = 111.2
myArray(17, 6) = 116.2
myArray(17, 7) = 360.7
myArray(17, 8) = 531.9
myArray(17, 9) = 290.2
myArray(17, 10) = 182.5
myArray(17, 11) = 34.5
myArray(17, 12) = 92.6
myArray(17, 13) = 14.6

myArray(18, 1) = 2012
myArray(18, 2) = 17.8
myArray(18, 3) = 3.7
myArray(18, 4) = 65.1
myArray(18, 5) = 106.8
myArray(18, 6) = 31.2
myArray(18, 7) = 93.7
myArray(18, 8) = 257.4
myArray(18, 9) = 479.5
myArray(18, 10) = 162.5
myArray(18, 11) = 61.2
myArray(18, 12) = 52.1
myArray(18, 13) = 56.6

myArray(19, 1) = 2013
myArray(19, 2) = 30.5
myArray(19, 3) = 33.2
myArray(19, 4) = 46.8
myArray(19, 5) = 65
myArray(19, 6) = 97.9
myArray(19, 7) = 229.9
myArray(19, 8) = 253.6
myArray(19, 9) = 183.9
myArray(19, 10) = 162.6
myArray(19, 11) = 25
myArray(19, 12) = 75
myArray(19, 13) = 37.3

myArray(20, 1) = 2014
myArray(20, 2) = 5.9
myArray(20, 3) = 6.8
myArray(20, 4) = 51.1
myArray(20, 5) = 43.7
myArray(20, 6) = 35
myArray(20, 7) = 92.6
myArray(20, 8) = 125.1
myArray(20, 9) = 197.5
myArray(20, 10) = 147.5
myArray(20, 11) = 151.1
myArray(20, 12) = 24.8
myArray(20, 13) = 32.6

myArray(21, 1) = 2015
myArray(21, 2) = 16
myArray(21, 3) = 26.5
myArray(21, 4) = 44.1
myArray(21, 5) = 109.1
myArray(21, 6) = 24.4
myArray(21, 7) = 83.3
myArray(21, 8) = 141.4
myArray(21, 9) = 54.3
myArray(21, 10) = 20.1
myArray(21, 11) = 90.5
myArray(21, 12) = 107.5
myArray(21, 13) = 39.7

myArray(22, 1) = 2016
myArray(22, 2) = 5.7
myArray(22, 3) = 45.5
myArray(22, 4) = 13.2
myArray(22, 5) = 132.1
myArray(22, 6) = 84.4
myArray(22, 7) = 39.9
myArray(22, 8) = 320
myArray(22, 9) = 69
myArray(22, 10) = 78.1
myArray(22, 11) = 83.6
myArray(22, 12) = 26.4
myArray(22, 13) = 40.1

myArray(23, 1) = 2017
myArray(23, 2) = 12
myArray(23, 3) = 38.7
myArray(23, 4) = 8.9
myArray(23, 5) = 61.7
myArray(23, 6) = 11.9
myArray(23, 7) = 17.5
myArray(23, 8) = 789.1
myArray(23, 9) = 225.2
myArray(23, 10) = 78.3
myArray(23, 11) = 23.1
myArray(23, 12) = 13.7
myArray(23, 13) = 21.1

myArray(24, 1) = 2018
myArray(24, 2) = 17.6
myArray(24, 3) = 30.6
myArray(24, 4) = 81.7
myArray(24, 5) = 133
myArray(24, 6) = 92
myArray(24, 7) = 63.3
myArray(24, 8) = 324.9
myArray(24, 9) = 247.9
myArray(24, 10) = 204
myArray(24, 11) = 112.2
myArray(24, 12) = 45.9
myArray(24, 13) = 28.5

myArray(25, 1) = 2019
myArray(25, 2) = 0.1
myArray(25, 3) = 23
myArray(25, 4) = 20.3
myArray(25, 5) = 60.8
myArray(25, 6) = 20.3
myArray(25, 7) = 82.5
myArray(25, 8) = 204.8
myArray(25, 9) = 80.5
myArray(25, 10) = 155.1
myArray(25, 11) = 84.3
myArray(25, 12) = 104.9
myArray(25, 13) = 20.1

myArray(26, 1) = 2020
myArray(26, 2) = 62
myArray(26, 3) = 62.7
myArray(26, 4) = 22.9
myArray(26, 5) = 15.7
myArray(26, 6) = 65.3
myArray(26, 7) = 145.9
myArray(26, 8) = 386.6
myArray(26, 9) = 385.8
myArray(26, 10) = 160.6
myArray(26, 11) = 5.8
myArray(26, 12) = 41
myArray(26, 13) = 4.3

myArray(27, 1) = 2021
myArray(27, 2) = 12.7
myArray(27, 3) = 7.5
myArray(27, 4) = 76.6
myArray(27, 5) = 46.4
myArray(27, 6) = 136.4
myArray(27, 7) = 75.4
myArray(27, 8) = 138.1
myArray(27, 9) = 233.1
myArray(27, 10) = 185
myArray(27, 11) = 29.4
myArray(27, 12) = 57.3
myArray(27, 13) = 3.7

myArray(28, 1) = 2022
myArray(28, 2) = 1.4
myArray(28, 3) = 2.4
myArray(28, 4) = 59
myArray(28, 5) = 45.2
myArray(28, 6) = 9.1
myArray(28, 7) = 129.6
myArray(28, 8) = 171.7
myArray(28, 9) = 519.4
myArray(28, 10) = 116
myArray(28, 11) = 105.9
myArray(28, 12) = 56.7
myArray(28, 13) = 20

myArray(29, 1) = 2023
myArray(29, 2) = 28
myArray(29, 3) = 2.8
myArray(29, 4) = 18.8
myArray(29, 5) = 30.1
myArray(29, 6) = 202.4
myArray(29, 7) = 100.5
myArray(29, 8) = 698.5
myArray(29, 9) = 297.7
myArray(29, 10) = 270.6
myArray(29, 11) = 17.4
myArray(29, 12) = 41.5
myArray(29, 13) = 97.3

myArray(30, 1) = 2024
myArray(30, 2) = 37.2
myArray(30, 3) = 79.1
myArray(30, 4) = 37.9
myArray(30, 5) = 49.1
myArray(30, 6) = 122.7
myArray(30, 7) = 84.7
myArray(30, 8) = 520.6
myArray(30, 9) = 113.1
myArray(30, 10) = 328
myArray(30, 11) = 112.7
myArray(30, 12) = 41.2
myArray(30, 13) = 8.6


    data_CHEONGJU = myArray

End Function


Function data_CHEONAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 19
myArray(1, 3) = 8.2
myArray(1, 4) = 25.3
myArray(1, 5) = 47
myArray(1, 6) = 48
myArray(1, 7) = 14.5
myArray(1, 8) = 239.9
myArray(1, 9) = 1082.5
myArray(1, 10) = 29
myArray(1, 11) = 23.5
myArray(1, 12) = 40.2
myArray(1, 13) = 8.9

myArray(2, 1) = 1996
myArray(2, 2) = 29.5
myArray(2, 3) = 10.2
myArray(2, 4) = 115
myArray(2, 5) = 54.5
myArray(2, 6) = 19
myArray(2, 7) = 237
myArray(2, 8) = 177.5
myArray(2, 9) = 116.5
myArray(2, 10) = 8
myArray(2, 11) = 102.5
myArray(2, 12) = 71.6
myArray(2, 13) = 26.2

myArray(3, 1) = 1997
myArray(3, 2) = 10.7
myArray(3, 3) = 44.1
myArray(3, 4) = 30
myArray(3, 5) = 66.5
myArray(3, 6) = 211
myArray(3, 7) = 191.5
myArray(3, 8) = 305
myArray(3, 9) = 175.5
myArray(3, 10) = 14.5
myArray(3, 11) = 23
myArray(3, 12) = 153.5
myArray(3, 13) = 43.5

myArray(4, 1) = 1998
myArray(4, 2) = 20.4
myArray(4, 3) = 27.9
myArray(4, 4) = 29.5
myArray(4, 5) = 120.5
myArray(4, 6) = 85
myArray(4, 7) = 219.5
myArray(4, 8) = 277
myArray(4, 9) = 408.1
myArray(4, 10) = 283
myArray(4, 11) = 51.5
myArray(4, 12) = 52.8
myArray(4, 13) = 8.5

myArray(5, 1) = 1999
myArray(5, 2) = 2.7
myArray(5, 3) = 2.8
myArray(5, 4) = 46.5
myArray(5, 5) = 88.5
myArray(5, 6) = 121.5
myArray(5, 7) = 163.7
myArray(5, 8) = 138.5
myArray(5, 9) = 313.5
myArray(5, 10) = 324.5
myArray(5, 11) = 134.5
myArray(5, 12) = 16.5
myArray(5, 13) = 11.9

myArray(6, 1) = 2000
myArray(6, 2) = 52.3
myArray(6, 3) = 2.7
myArray(6, 4) = 7.1
myArray(6, 5) = 36
myArray(6, 6) = 36
myArray(6, 7) = 181
myArray(6, 8) = 83
myArray(6, 9) = 636
myArray(6, 10) = 287.5
myArray(6, 11) = 32
myArray(6, 12) = 32
myArray(6, 13) = 22.5

myArray(7, 1) = 2001
myArray(7, 2) = 43.5
myArray(7, 3) = 44
myArray(7, 4) = 16.5
myArray(7, 5) = 19
myArray(7, 6) = 15
myArray(7, 7) = 227.5
myArray(7, 8) = 178
myArray(7, 9) = 194.5
myArray(7, 10) = 12
myArray(7, 11) = 63.5
myArray(7, 12) = 6.3
myArray(7, 13) = 18.4

myArray(8, 1) = 2002
myArray(8, 2) = 45.3
myArray(8, 3) = 6
myArray(8, 4) = 25.5
myArray(8, 5) = 128
myArray(8, 6) = 104
myArray(8, 7) = 54
myArray(8, 8) = 229.5
myArray(8, 9) = 481.5
myArray(8, 10) = 57
myArray(8, 11) = 91.5
myArray(8, 12) = 42.1
myArray(8, 13) = 48.1

myArray(9, 1) = 2003
myArray(9, 2) = 18.6
myArray(9, 3) = 44
myArray(9, 4) = 38.1
myArray(9, 5) = 172.3
myArray(9, 6) = 106
myArray(9, 7) = 178.6
myArray(9, 8) = 381.2
myArray(9, 9) = 334.6
myArray(9, 10) = 264.2
myArray(9, 11) = 27
myArray(9, 12) = 46.7
myArray(9, 13) = 17

myArray(10, 1) = 2004
myArray(10, 2) = 16.4
myArray(10, 3) = 21.3
myArray(10, 4) = 21.5
myArray(10, 5) = 67.5
myArray(10, 6) = 127.6
myArray(10, 7) = 235
myArray(10, 8) = 365.2
myArray(10, 9) = 229.3
myArray(10, 10) = 189
myArray(10, 11) = 4.5
myArray(10, 12) = 53
myArray(10, 13) = 33

myArray(11, 1) = 2005
myArray(11, 2) = 3
myArray(11, 3) = 29.8
myArray(11, 4) = 37
myArray(11, 5) = 53.7
myArray(11, 6) = 48
myArray(11, 7) = 183
myArray(11, 8) = 313.8
myArray(11, 9) = 202
myArray(11, 10) = 377.5
myArray(11, 11) = 26.7
myArray(11, 12) = 23.5
myArray(11, 13) = 11.3

myArray(12, 1) = 2006
myArray(12, 2) = 25.2
myArray(12, 3) = 18.5
myArray(12, 4) = 6.1
myArray(12, 5) = 78.6
myArray(12, 6) = 79
myArray(12, 7) = 120
myArray(12, 8) = 535.1
myArray(12, 9) = 63.5
myArray(12, 10) = 22.2
myArray(12, 11) = 21.6
myArray(12, 12) = 56.3
myArray(12, 13) = 17.2

myArray(13, 1) = 2007
myArray(13, 2) = 9.4
myArray(13, 3) = 34.1
myArray(13, 4) = 108.3
myArray(13, 5) = 35.3
myArray(13, 6) = 126.2
myArray(13, 7) = 106.7
myArray(13, 8) = 215.6
myArray(13, 9) = 470.6
myArray(13, 10) = 353.3
myArray(13, 11) = 43.4
myArray(13, 12) = 15.6
myArray(13, 13) = 43.9

myArray(14, 1) = 2008
myArray(14, 2) = 17.5
myArray(14, 3) = 11.1
myArray(14, 4) = 40.1
myArray(14, 5) = 34
myArray(14, 6) = 62.6
myArray(14, 7) = 126.7
myArray(14, 8) = 287.2
myArray(14, 9) = 138.8
myArray(14, 10) = 89.3
myArray(14, 11) = 30.4
myArray(14, 12) = 16.6
myArray(14, 13) = 15.8

myArray(15, 1) = 2009
myArray(15, 2) = 13.3
myArray(15, 3) = 16
myArray(15, 4) = 51.6
myArray(15, 5) = 30.6
myArray(15, 6) = 112.6
myArray(15, 7) = 55.6
myArray(15, 8) = 335.8
myArray(15, 9) = 212.3
myArray(15, 10) = 30.8
myArray(15, 11) = 61.1
myArray(15, 12) = 39.7
myArray(15, 13) = 40.5

myArray(16, 1) = 2010
myArray(16, 2) = 40.7
myArray(16, 3) = 50.4
myArray(16, 4) = 73.8
myArray(16, 5) = 61
myArray(16, 6) = 84
myArray(16, 7) = 37
myArray(16, 8) = 171
myArray(16, 9) = 486.1
myArray(16, 10) = 316.9
myArray(16, 11) = 19.4
myArray(16, 12) = 13.5
myArray(16, 13) = 24.5

myArray(17, 1) = 2011
myArray(17, 2) = 7.9
myArray(17, 3) = 31
myArray(17, 4) = 26.5
myArray(17, 5) = 133.2
myArray(17, 6) = 103.3
myArray(17, 7) = 374.6
myArray(17, 8) = 645.1
myArray(17, 9) = 268.2
myArray(17, 10) = 153.2
myArray(17, 11) = 26.5
myArray(17, 12) = 65.8
myArray(17, 13) = 10.5

myArray(18, 1) = 2012
myArray(18, 2) = 14.5
myArray(18, 3) = 2.3
myArray(18, 4) = 44.9
myArray(18, 5) = 81.6
myArray(18, 6) = 16.8
myArray(18, 7) = 75.1
myArray(18, 8) = 252.5
myArray(18, 9) = 483.7
myArray(18, 10) = 190.1
myArray(18, 11) = 66.6
myArray(18, 12) = 52.6
myArray(18, 13) = 56

myArray(19, 1) = 2013
myArray(19, 2) = 28.5
myArray(19, 3) = 35.2
myArray(19, 4) = 40
myArray(19, 5) = 56.3
myArray(19, 6) = 123.5
myArray(19, 7) = 102.1
myArray(19, 8) = 308.2
myArray(19, 9) = 173.6
myArray(19, 10) = 117.5
myArray(19, 11) = 12.2
myArray(19, 12) = 58.2
myArray(19, 13) = 40.3

myArray(20, 1) = 2014
myArray(20, 2) = 4.9
myArray(20, 3) = 14.7
myArray(20, 4) = 40.9
myArray(20, 5) = 62.1
myArray(20, 6) = 34.6
myArray(20, 7) = 73.9
myArray(20, 8) = 239
myArray(20, 9) = 218.7
myArray(20, 10) = 144
myArray(20, 11) = 119.5
myArray(20, 12) = 28.9
myArray(20, 13) = 38.9

myArray(21, 1) = 2015
myArray(21, 2) = 12.7
myArray(21, 3) = 21.5
myArray(21, 4) = 23.3
myArray(21, 5) = 87.6
myArray(21, 6) = 27.5
myArray(21, 7) = 86
myArray(21, 8) = 136.8
myArray(21, 9) = 64.2
myArray(21, 10) = 29
myArray(21, 11) = 69
myArray(21, 12) = 128.6
myArray(21, 13) = 41.8

myArray(22, 1) = 2016
myArray(22, 2) = 8
myArray(22, 3) = 43.6
myArray(22, 4) = 16.5
myArray(22, 5) = 118.3
myArray(22, 6) = 107.2
myArray(22, 7) = 36.2
myArray(22, 8) = 364.3
myArray(22, 9) = 82
myArray(22, 10) = 55
myArray(22, 11) = 95.9
myArray(22, 12) = 33.5
myArray(22, 13) = 44.3

myArray(23, 1) = 2017
myArray(23, 2) = 13.9
myArray(23, 3) = 32.2
myArray(23, 4) = 6.5
myArray(23, 5) = 42.9
myArray(23, 6) = 14.3
myArray(23, 7) = 15.6
myArray(23, 8) = 788.1
myArray(23, 9) = 291.5
myArray(23, 10) = 43.3
myArray(23, 11) = 14.1
myArray(23, 12) = 23.8
myArray(23, 13) = 18.8

myArray(24, 1) = 2018
myArray(24, 2) = 14
myArray(24, 3) = 31.6
myArray(24, 4) = 62.2
myArray(24, 5) = 117
myArray(24, 6) = 82.7
myArray(24, 7) = 88.9
myArray(24, 8) = 185.8
myArray(24, 9) = 282.7
myArray(24, 10) = 124.6
myArray(24, 11) = 99.8
myArray(24, 12) = 48.3
myArray(24, 13) = 25.8

myArray(25, 1) = 2019
myArray(25, 2) = 0.6
myArray(25, 3) = 25.5
myArray(25, 4) = 26.9
myArray(25, 5) = 43.9
myArray(25, 6) = 15.1
myArray(25, 7) = 84.9
myArray(25, 8) = 234.7
myArray(25, 9) = 90.7
myArray(25, 10) = 102.8
myArray(25, 11) = 81.9
myArray(25, 12) = 120.6
myArray(25, 13) = 18

myArray(26, 1) = 2020
myArray(26, 2) = 59.7
myArray(26, 3) = 63.1
myArray(26, 4) = 21.7
myArray(26, 5) = 15.1
myArray(26, 6) = 86.4
myArray(26, 7) = 121.9
myArray(26, 8) = 356.4
myArray(26, 9) = 481.7
myArray(26, 10) = 167.2
myArray(26, 11) = 18.9
myArray(26, 12) = 45.9
myArray(26, 13) = 5.5

myArray(27, 1) = 2021
myArray(27, 2) = 17.8
myArray(27, 3) = 9.2
myArray(27, 4) = 75.3
myArray(27, 5) = 54.7
myArray(27, 6) = 135.9
myArray(27, 7) = 44.8
myArray(27, 8) = 117.6
myArray(27, 9) = 230
myArray(27, 10) = 250.8
myArray(27, 11) = 49.5
myArray(27, 12) = 67.9
myArray(27, 13) = 5.4

myArray(28, 1) = 2022
myArray(28, 2) = 3.3
myArray(28, 3) = 3.3
myArray(28, 4) = 57.6
myArray(28, 5) = 51.6
myArray(28, 6) = 9.8
myArray(28, 7) = 168
myArray(28, 8) = 117
myArray(28, 9) = 366.6
myArray(28, 10) = 133.3
myArray(28, 11) = 98.2
myArray(28, 12) = 43.2
myArray(28, 13) = 28.8

myArray(29, 1) = 2023
myArray(29, 2) = 31
myArray(29, 3) = 3.1
myArray(29, 4) = 16.4
myArray(29, 5) = 29.6
myArray(29, 6) = 116.9
myArray(29, 7) = 178.9
myArray(29, 8) = 574.9
myArray(29, 9) = 196.5
myArray(29, 10) = 180.1
myArray(29, 11) = 28.7
myArray(29, 12) = 56.9
myArray(29, 13) = 89.5

myArray(30, 1) = 2024
myArray(30, 2) = 34.9
myArray(30, 3) = 89.7
myArray(30, 4) = 41.3
myArray(30, 5) = 48.2
myArray(30, 6) = 107.7
myArray(30, 7) = 106.4
myArray(30, 8) = 509.9
myArray(30, 9) = 66.4
myArray(30, 10) = 318
myArray(30, 11) = 99.7
myArray(30, 12) = 39.5
myArray(30, 13) = 15.2


    data_CHEONAN = myArray

End Function


Function data_JAECHEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 14
myArray(1, 3) = 6.1
myArray(1, 4) = 43.3
myArray(1, 5) = 60
myArray(1, 6) = 64.5
myArray(1, 7) = 72.5
myArray(1, 8) = 292.5
myArray(1, 9) = 742.5
myArray(1, 10) = 66
myArray(1, 11) = 38.5
myArray(1, 12) = 42
myArray(1, 13) = 6

myArray(2, 1) = 1996
myArray(2, 2) = 33.2
myArray(2, 3) = 8.8
myArray(2, 4) = 111.5
myArray(2, 5) = 46.5
myArray(2, 6) = 33.5
myArray(2, 7) = 174
myArray(2, 8) = 264
myArray(2, 9) = 121.5
myArray(2, 10) = 23
myArray(2, 11) = 90
myArray(2, 12) = 62.4
myArray(2, 13) = 18.8

myArray(3, 1) = 1997
myArray(3, 2) = 9.3
myArray(3, 3) = 48.4
myArray(3, 4) = 29.5
myArray(3, 5) = 65
myArray(3, 6) = 203
myArray(3, 7) = 151
myArray(3, 8) = 423.5
myArray(3, 9) = 197.3
myArray(3, 10) = 85.5
myArray(3, 11) = 16.2
myArray(3, 12) = 123.2
myArray(3, 13) = 32.9

myArray(4, 1) = 1998
myArray(4, 2) = 15.4
myArray(4, 3) = 27.1
myArray(4, 4) = 27
myArray(4, 5) = 136.5
myArray(4, 6) = 101
myArray(4, 7) = 195.8
myArray(4, 8) = 289.5
myArray(4, 9) = 546.1
myArray(4, 10) = 132.2
myArray(4, 11) = 72.5
myArray(4, 12) = 34.7
myArray(4, 13) = 3.6

myArray(5, 1) = 1999
myArray(5, 2) = 4.9
myArray(5, 3) = 4.2
myArray(5, 4) = 66.1
myArray(5, 5) = 129.5
myArray(5, 6) = 104
myArray(5, 7) = 136.5
myArray(5, 8) = 214
myArray(5, 9) = 323
myArray(5, 10) = 281.5
myArray(5, 11) = 159
myArray(5, 12) = 23.5
myArray(5, 13) = 7.2

myArray(6, 1) = 2000
myArray(6, 2) = 39.9
myArray(6, 3) = 4.1
myArray(6, 4) = 15.5
myArray(6, 5) = 41.5
myArray(6, 6) = 96
myArray(6, 7) = 196.5
myArray(6, 8) = 197.2
myArray(6, 9) = 259
myArray(6, 10) = 221.5
myArray(6, 11) = 24
myArray(6, 12) = 34
myArray(6, 13) = 19.9

myArray(7, 1) = 2001
myArray(7, 2) = 46.4
myArray(7, 3) = 49.6
myArray(7, 4) = 18.8
myArray(7, 5) = 18.5
myArray(7, 6) = 9
myArray(7, 7) = 269.5
myArray(7, 8) = 227.5
myArray(7, 9) = 93.5
myArray(7, 10) = 19.5
myArray(7, 11) = 84
myArray(7, 12) = 3
myArray(7, 13) = 10

myArray(8, 1) = 2002
myArray(8, 2) = 42.3
myArray(8, 3) = 3.2
myArray(8, 4) = 24
myArray(8, 5) = 199
myArray(8, 6) = 92
myArray(8, 7) = 89
myArray(8, 8) = 214.5
myArray(8, 9) = 652
myArray(8, 10) = 69.5
myArray(8, 11) = 44.5
myArray(8, 12) = 13
myArray(8, 13) = 57.4

myArray(9, 1) = 2003
myArray(9, 2) = 11.2
myArray(9, 3) = 48.7
myArray(9, 4) = 42
myArray(9, 5) = 198
myArray(9, 6) = 149.5
myArray(9, 7) = 196.5
myArray(9, 8) = 495
myArray(9, 9) = 346
myArray(9, 10) = 287.5
myArray(9, 11) = 24.5
myArray(9, 12) = 60.5
myArray(9, 13) = 17.2

myArray(10, 1) = 2004
myArray(10, 2) = 15.2
myArray(10, 3) = 31.5
myArray(10, 4) = 34.5
myArray(10, 5) = 74
myArray(10, 6) = 142
myArray(10, 7) = 395.5
myArray(10, 8) = 455.5
myArray(10, 9) = 259
myArray(10, 10) = 163.5
myArray(10, 11) = 2
myArray(10, 12) = 37
myArray(10, 13) = 21.1

myArray(11, 1) = 2005
myArray(11, 2) = 3.8
myArray(11, 3) = 17.9
myArray(11, 4) = 44
myArray(11, 5) = 77.5
myArray(11, 6) = 78.5
myArray(11, 7) = 171.5
myArray(11, 8) = 424.5
myArray(11, 9) = 259
myArray(11, 10) = 363
myArray(11, 11) = 57.5
myArray(11, 12) = 19.5
myArray(11, 13) = 8.5

myArray(12, 1) = 2006
myArray(12, 2) = 26.5
myArray(12, 3) = 33.3
myArray(12, 4) = 12.1
myArray(12, 5) = 106
myArray(12, 6) = 105.5
myArray(12, 7) = 145
myArray(12, 8) = 1111
myArray(12, 9) = 56.5
myArray(12, 10) = 25
myArray(12, 11) = 39.5
myArray(12, 12) = 46.3
myArray(12, 13) = 13.1

myArray(13, 1) = 2007
myArray(13, 2) = 11
myArray(13, 3) = 34
myArray(13, 4) = 177.2
myArray(13, 5) = 19.5
myArray(13, 6) = 151.5
myArray(13, 7) = 124
myArray(13, 8) = 442.5
myArray(13, 9) = 696.5
myArray(13, 10) = 333
myArray(13, 11) = 33.5
myArray(13, 12) = 24.3
myArray(13, 13) = 20.3

myArray(14, 1) = 2008
myArray(14, 2) = 19.8
myArray(14, 3) = 6.5
myArray(14, 4) = 63.8
myArray(14, 5) = 41.5
myArray(14, 6) = 54
myArray(14, 7) = 86.5
myArray(14, 8) = 274.6
myArray(14, 9) = 222
myArray(14, 10) = 63.6
myArray(14, 11) = 24
myArray(14, 12) = 8.4
myArray(14, 13) = 21.1

myArray(15, 1) = 2009
myArray(15, 2) = 14.8
myArray(15, 3) = 27.5
myArray(15, 4) = 57.1
myArray(15, 5) = 40.4
myArray(15, 6) = 121.7
myArray(15, 7) = 162.1
myArray(15, 8) = 474.5
myArray(15, 9) = 213
myArray(15, 10) = 58.5
myArray(15, 11) = 31.5
myArray(15, 12) = 44.3
myArray(15, 13) = 32

myArray(16, 1) = 2010
myArray(16, 2) = 61.1
myArray(16, 3) = 68.6
myArray(16, 4) = 117.3
myArray(16, 5) = 71.5
myArray(16, 6) = 118.1
myArray(16, 7) = 87.6
myArray(16, 8) = 180.1
myArray(16, 9) = 345.7
myArray(16, 10) = 432.8
myArray(16, 11) = 22.8
myArray(16, 12) = 22.4
myArray(16, 13) = 17.2

myArray(17, 1) = 2011
myArray(17, 2) = 2.8
myArray(17, 3) = 52.9
myArray(17, 4) = 36.5
myArray(17, 5) = 189.8
myArray(17, 6) = 121.8
myArray(17, 7) = 459
myArray(17, 8) = 665.2
myArray(17, 9) = 398.5
myArray(17, 10) = 157.2
myArray(17, 11) = 55.4
myArray(17, 12) = 81.1
myArray(17, 13) = 10.3

myArray(18, 1) = 2012
myArray(18, 2) = 16
myArray(18, 3) = 5
myArray(18, 4) = 83.2
myArray(18, 5) = 135.5
myArray(18, 6) = 40.9
myArray(18, 7) = 108.1
myArray(18, 8) = 344.8
myArray(18, 9) = 319.9
myArray(18, 10) = 144.5
myArray(18, 11) = 67.1
myArray(18, 12) = 68.7
myArray(18, 13) = 47.6

myArray(19, 1) = 2013
myArray(19, 2) = 40.5
myArray(19, 3) = 55
myArray(19, 4) = 48
myArray(19, 5) = 92.3
myArray(19, 6) = 118.5
myArray(19, 7) = 144.6
myArray(19, 8) = 442.4
myArray(19, 9) = 274.3
myArray(19, 10) = 118.9
myArray(19, 11) = 8.9
myArray(19, 12) = 63.2
myArray(19, 13) = 30.5

myArray(20, 1) = 2014
myArray(20, 2) = 10.5
myArray(20, 3) = 23.6
myArray(20, 4) = 44.5
myArray(20, 5) = 49.5
myArray(20, 6) = 41.4
myArray(20, 7) = 62.1
myArray(20, 8) = 111.4
myArray(20, 9) = 246.2
myArray(20, 10) = 131
myArray(20, 11) = 150.6
myArray(20, 12) = 24
myArray(20, 13) = 18.8

myArray(21, 1) = 2015
myArray(21, 2) = 17.5
myArray(21, 3) = 32.2
myArray(21, 4) = 31.7
myArray(21, 5) = 83.5
myArray(21, 6) = 31.5
myArray(21, 7) = 75.4
myArray(21, 8) = 225.1
myArray(21, 9) = 63.8
myArray(21, 10) = 36.6
myArray(21, 11) = 68.1
myArray(21, 12) = 110.6
myArray(21, 13) = 27.4

myArray(22, 1) = 2016
myArray(22, 2) = 4.5
myArray(22, 3) = 68.7
myArray(22, 4) = 21.5
myArray(22, 5) = 117.1
myArray(22, 6) = 82.4
myArray(22, 7) = 42.1
myArray(22, 8) = 419.7
myArray(22, 9) = 113.8
myArray(22, 10) = 47.3
myArray(22, 11) = 109.6
myArray(22, 12) = 22.1
myArray(22, 13) = 59.1

myArray(23, 1) = 2017
myArray(23, 2) = 10.2
myArray(23, 3) = 29.5
myArray(23, 4) = 24.3
myArray(23, 5) = 70.8
myArray(23, 6) = 12.5
myArray(23, 7) = 69.6
myArray(23, 8) = 464.8
myArray(23, 9) = 265.1
myArray(23, 10) = 43.3
myArray(23, 11) = 22.5
myArray(23, 12) = 34.1
myArray(23, 13) = 24

myArray(24, 1) = 2018
myArray(24, 2) = 6.9
myArray(24, 3) = 23.8
myArray(24, 4) = 61.7
myArray(24, 5) = 112.9
myArray(24, 6) = 172.5
myArray(24, 7) = 138.5
myArray(24, 8) = 161.5
myArray(24, 9) = 350.3
myArray(24, 10) = 185.5
myArray(24, 11) = 105.4
myArray(24, 12) = 60.3
myArray(24, 13) = 30

myArray(25, 1) = 2019
myArray(25, 2) = 1
myArray(25, 3) = 28.5
myArray(25, 4) = 49.5
myArray(25, 5) = 58.1
myArray(25, 6) = 26.1
myArray(25, 7) = 90
myArray(25, 8) = 158.6
myArray(25, 9) = 99.1
myArray(25, 10) = 164.5
myArray(25, 11) = 66.9
myArray(25, 12) = 81
myArray(25, 13) = 19.7

myArray(26, 1) = 2020
myArray(26, 2) = 68.1
myArray(26, 3) = 57.4
myArray(26, 4) = 14.4
myArray(26, 5) = 28.9
myArray(26, 6) = 130.1
myArray(26, 7) = 78.6
myArray(26, 8) = 317.5
myArray(26, 9) = 685.8
myArray(26, 10) = 134.1
myArray(26, 11) = 12.4
myArray(26, 12) = 13.4
myArray(26, 13) = 11.1

myArray(27, 1) = 2021
myArray(27, 2) = 18.2
myArray(27, 3) = 12.2
myArray(27, 4) = 100.1
myArray(27, 5) = 104
myArray(27, 6) = 157.8
myArray(27, 7) = 81.6
myArray(27, 8) = 223.1
myArray(27, 9) = 174.9
myArray(27, 10) = 209.2
myArray(27, 11) = 26.6
myArray(27, 12) = 48
myArray(27, 13) = 6.5

myArray(28, 1) = 2022
myArray(28, 2) = 3
myArray(28, 3) = 6.8
myArray(28, 4) = 88.3
myArray(28, 5) = 40.5
myArray(28, 6) = 7.5
myArray(28, 7) = 273.9
myArray(28, 8) = 310.9
myArray(28, 9) = 485.3
myArray(28, 10) = 93.7
myArray(28, 11) = 87.1
myArray(28, 12) = 60
myArray(28, 13) = 19.6

myArray(29, 1) = 2023
myArray(29, 2) = 32.6
myArray(29, 3) = 2.2
myArray(29, 4) = 14
myArray(29, 5) = 48.6
myArray(29, 6) = 167.1
myArray(29, 7) = 231.3
myArray(29, 8) = 605.3
myArray(29, 9) = 240.9
myArray(29, 10) = 234.6
myArray(29, 11) = 33.8
myArray(29, 12) = 49.8
myArray(29, 13) = 105.5

myArray(30, 1) = 2024
myArray(30, 2) = 24.8
myArray(30, 3) = 78.9
myArray(30, 4) = 52.2
myArray(30, 5) = 54.7
myArray(30, 6) = 128.7
myArray(30, 7) = 171.7
myArray(30, 8) = 467.8
myArray(30, 9) = 99.1
myArray(30, 10) = 195.6
myArray(30, 11) = 114.1
myArray(30, 12) = 42.1
myArray(30, 13) = 3.1


    data_JAECHEON = myArray

End Function


Function data_SEOSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 22.7
myArray(1, 3) = 7.2
myArray(1, 4) = 37.3
myArray(1, 5) = 48.2
myArray(1, 6) = 67.1
myArray(1, 7) = 24.5
myArray(1, 8) = 144.1
myArray(1, 9) = 992.7
myArray(1, 10) = 20.2
myArray(1, 11) = 19.3
myArray(1, 12) = 49.9
myArray(1, 13) = 15.1

myArray(2, 1) = 1996
myArray(2, 2) = 29.1
myArray(2, 3) = 5.7
myArray(2, 4) = 115.1
myArray(2, 5) = 48.1
myArray(2, 6) = 20
myArray(2, 7) = 179.2
myArray(2, 8) = 152.8
myArray(2, 9) = 74.1
myArray(2, 10) = 6.4
myArray(2, 11) = 92.2
myArray(2, 12) = 72.1
myArray(2, 13) = 35.3

myArray(3, 1) = 1997
myArray(3, 2) = 20.5
myArray(3, 3) = 32.5
myArray(3, 4) = 29.6
myArray(3, 5) = 69.5
myArray(3, 6) = 232.8
myArray(3, 7) = 204.4
myArray(3, 8) = 298.7
myArray(3, 9) = 87.2
myArray(3, 10) = 16.1
myArray(3, 11) = 8.7
myArray(3, 12) = 116.7
myArray(3, 13) = 40.2

myArray(4, 1) = 1998
myArray(4, 2) = 40.1
myArray(4, 3) = 54.2
myArray(4, 4) = 35
myArray(4, 5) = 160.6
myArray(4, 6) = 95.5
myArray(4, 7) = 281.7
myArray(4, 8) = 295.6
myArray(4, 9) = 491.8
myArray(4, 10) = 168
myArray(4, 11) = 24.3
myArray(4, 12) = 55.6
myArray(4, 13) = 9.2

myArray(5, 1) = 1999
myArray(5, 2) = 8
myArray(5, 3) = 7.8
myArray(5, 4) = 59.9
myArray(5, 5) = 90.1
myArray(5, 6) = 178.8
myArray(5, 7) = 105.1
myArray(5, 8) = 175.6
myArray(5, 9) = 497.4
myArray(5, 10) = 532.6
myArray(5, 11) = 111.3
myArray(5, 12) = 36.6
myArray(5, 13) = 23.4

myArray(6, 1) = 2000
myArray(6, 2) = 63
myArray(6, 3) = 2.9
myArray(6, 4) = 3.7
myArray(6, 5) = 38.1
myArray(6, 6) = 62.1
myArray(6, 7) = 204.4
myArray(6, 8) = 60.8
myArray(6, 9) = 608.1
myArray(6, 10) = 298.1
myArray(6, 11) = 34.4
myArray(6, 12) = 24.8
myArray(6, 13) = 24.4

myArray(7, 1) = 2001
myArray(7, 2) = 66.9
myArray(7, 3) = 40.4
myArray(7, 4) = 12.7
myArray(7, 5) = 18.7
myArray(7, 6) = 17.8
myArray(7, 7) = 200.2
myArray(7, 8) = 402
myArray(7, 9) = 136.6
myArray(7, 10) = 15
myArray(7, 11) = 47.5
myArray(7, 12) = 8.2
myArray(7, 13) = 20.8

myArray(8, 1) = 2002
myArray(8, 2) = 22.5
myArray(8, 3) = 7
myArray(8, 4) = 29.3
myArray(8, 5) = 179.5
myArray(8, 6) = 177.3
myArray(8, 7) = 60.8
myArray(8, 8) = 296.1
myArray(8, 9) = 428.2
myArray(8, 10) = 57.5
myArray(8, 11) = 78.3
myArray(8, 12) = 36.3
myArray(8, 13) = 14.8

myArray(9, 1) = 2003
myArray(9, 2) = 20.9
myArray(9, 3) = 39
myArray(9, 4) = 22.5
myArray(9, 5) = 180
myArray(9, 6) = 105.5
myArray(9, 7) = 221.8
myArray(9, 8) = 290.2
myArray(9, 9) = 257.9
myArray(9, 10) = 201.9
myArray(9, 11) = 23
myArray(9, 12) = 53.6
myArray(9, 13) = 17.1

myArray(10, 1) = 2004
myArray(10, 2) = 27.3
myArray(10, 3) = 26.3
myArray(10, 4) = 15.7
myArray(10, 5) = 80.2
myArray(10, 6) = 140.3
myArray(10, 7) = 211.1
myArray(10, 8) = 321.9
myArray(10, 9) = 131.2
myArray(10, 10) = 282.6
myArray(10, 11) = 1.8
myArray(10, 12) = 70.5
myArray(10, 13) = 32

myArray(11, 1) = 2005
myArray(11, 2) = 10.4
myArray(11, 3) = 34
myArray(11, 4) = 36.1
myArray(11, 5) = 77.2
myArray(11, 6) = 56.1
myArray(11, 7) = 147
myArray(11, 8) = 386.1
myArray(11, 9) = 270.5
myArray(11, 10) = 228.7
myArray(11, 11) = 30.9
myArray(11, 12) = 19.6
myArray(11, 13) = 37.6

myArray(12, 1) = 2006
myArray(12, 2) = 29.7
myArray(12, 3) = 18.9
myArray(12, 4) = 5
myArray(12, 5) = 77.3
myArray(12, 6) = 133.5
myArray(12, 7) = 226.8
myArray(12, 8) = 494.5
myArray(12, 9) = 58.2
myArray(12, 10) = 10.1
myArray(12, 11) = 10.5
myArray(12, 12) = 55
myArray(12, 13) = 19.7

myArray(13, 1) = 2007
myArray(13, 2) = 13
myArray(13, 3) = 25.5
myArray(13, 4) = 127.2
myArray(13, 5) = 28.1
myArray(13, 6) = 108.5
myArray(13, 7) = 123.5
myArray(13, 8) = 257
myArray(13, 9) = 414.6
myArray(13, 10) = 305.8
myArray(13, 11) = 30.7
myArray(13, 12) = 14.4
myArray(13, 13) = 22.8

myArray(14, 1) = 2008
myArray(14, 2) = 15
myArray(14, 3) = 7
myArray(14, 4) = 26
myArray(14, 5) = 46.1
myArray(14, 6) = 88.5
myArray(14, 7) = 118.1
myArray(14, 8) = 335.5
myArray(14, 9) = 114.2
myArray(14, 10) = 62.7
myArray(14, 11) = 34
myArray(14, 12) = 34.6
myArray(14, 13) = 27.9

myArray(15, 1) = 2009
myArray(15, 2) = 15.2
myArray(15, 3) = 26.5
myArray(15, 4) = 67
myArray(15, 5) = 43
myArray(15, 6) = 117.9
myArray(15, 7) = 74.9
myArray(15, 8) = 364.9
myArray(15, 9) = 196.3
myArray(15, 10) = 16
myArray(15, 11) = 49.2
myArray(15, 12) = 59.1
myArray(15, 13) = 44.3

myArray(16, 1) = 2010
myArray(16, 2) = 55.5
myArray(16, 3) = 58.4
myArray(16, 4) = 79.2
myArray(16, 5) = 52.2
myArray(16, 6) = 168
myArray(16, 7) = 94.9
myArray(16, 8) = 447.1
myArray(16, 9) = 707
myArray(16, 10) = 402
myArray(16, 11) = 29.1
myArray(16, 12) = 12
myArray(16, 13) = 36.4

myArray(17, 1) = 2011
myArray(17, 2) = 8.8
myArray(17, 3) = 55.8
myArray(17, 4) = 34.5
myArray(17, 5) = 96.2
myArray(17, 6) = 107.9
myArray(17, 7) = 462.6
myArray(17, 8) = 656.5
myArray(17, 9) = 151.2
myArray(17, 10) = 50.3
myArray(17, 11) = 18.1
myArray(17, 12) = 48.9
myArray(17, 13) = 13.6

myArray(18, 1) = 2012
myArray(18, 2) = 15.1
myArray(18, 3) = 2.4
myArray(18, 4) = 41.6
myArray(18, 5) = 113.5
myArray(18, 6) = 14.5
myArray(18, 7) = 91.1
myArray(18, 8) = 266.8
myArray(18, 9) = 647.9
myArray(18, 10) = 201.5
myArray(18, 11) = 100.7
myArray(18, 12) = 82.1
myArray(18, 13) = 65.4

myArray(19, 1) = 2013
myArray(19, 2) = 36.8
myArray(19, 3) = 64.8
myArray(19, 4) = 60.8
myArray(19, 5) = 61.8
myArray(19, 6) = 114.9
myArray(19, 7) = 94.4
myArray(19, 8) = 213.8
myArray(19, 9) = 120.6
myArray(19, 10) = 147.4
myArray(19, 11) = 5.7
myArray(19, 12) = 64.9
myArray(19, 13) = 32.8

myArray(20, 1) = 2014
myArray(20, 2) = 7
myArray(20, 3) = 17
myArray(20, 4) = 31.2
myArray(20, 5) = 85.6
myArray(20, 6) = 52.7
myArray(20, 7) = 69.3
myArray(20, 8) = 151.7
myArray(20, 9) = 242.3
myArray(20, 10) = 106.7
myArray(20, 11) = 117.2
myArray(20, 12) = 37.8
myArray(20, 13) = 81.6

myArray(21, 1) = 2015
myArray(21, 2) = 20.7
myArray(21, 3) = 23.1
myArray(21, 4) = 20.6
myArray(21, 5) = 116.8
myArray(21, 6) = 40.6
myArray(21, 7) = 64.1
myArray(21, 8) = 158.5
myArray(21, 9) = 63.1
myArray(21, 10) = 15.1
myArray(21, 11) = 73.1
myArray(21, 12) = 156.6
myArray(21, 13) = 63.6

myArray(22, 1) = 2016
myArray(22, 2) = 21.9
myArray(22, 3) = 61.7
myArray(22, 4) = 24.3
myArray(22, 5) = 87
myArray(22, 6) = 153.7
myArray(22, 7) = 36.8
myArray(22, 8) = 295.6
myArray(22, 9) = 34
myArray(22, 10) = 53.1
myArray(22, 11) = 73.8
myArray(22, 12) = 17.5
myArray(22, 13) = 62.7

myArray(23, 1) = 2017
myArray(23, 2) = 21.3
myArray(23, 3) = 31.4
myArray(23, 4) = 4.8
myArray(23, 5) = 38.9
myArray(23, 6) = 27.9
myArray(23, 7) = 23.3
myArray(23, 8) = 327.8
myArray(23, 9) = 231.3
myArray(23, 10) = 37.6
myArray(23, 11) = 25.5
myArray(23, 12) = 24.7
myArray(23, 13) = 35.9

myArray(24, 1) = 2018
myArray(24, 2) = 21
myArray(24, 3) = 40.5
myArray(24, 4) = 76.6
myArray(24, 5) = 132.8
myArray(24, 6) = 147.7
myArray(24, 7) = 162.3
myArray(24, 8) = 152.9
myArray(24, 9) = 156.8
myArray(24, 10) = 82.7
myArray(24, 11) = 153.2
myArray(24, 12) = 73.9
myArray(24, 13) = 26.8

myArray(25, 1) = 2019
myArray(25, 2) = 1.1
myArray(25, 3) = 30.2
myArray(25, 4) = 43.7
myArray(25, 5) = 43.9
myArray(25, 6) = 20.1
myArray(25, 7) = 56.3
myArray(25, 8) = 174.5
myArray(25, 9) = 121.1
myArray(25, 10) = 181.1
myArray(25, 11) = 81
myArray(25, 12) = 124.6
myArray(25, 13) = 37.4

myArray(26, 1) = 2020
myArray(26, 2) = 46
myArray(26, 3) = 72.3
myArray(26, 4) = 23
myArray(26, 5) = 20.7
myArray(26, 6) = 101.3
myArray(26, 7) = 144
myArray(26, 8) = 329.4
myArray(26, 9) = 400
myArray(26, 10) = 257.7
myArray(26, 11) = 12.6
myArray(26, 12) = 72
myArray(26, 13) = 9.7

myArray(27, 1) = 2021
myArray(27, 2) = 25.3
myArray(27, 3) = 9.6
myArray(27, 4) = 112.8
myArray(27, 5) = 110.6
myArray(27, 6) = 132.3
myArray(27, 7) = 70.9
myArray(27, 8) = 121.6
myArray(27, 9) = 217.8
myArray(27, 10) = 206
myArray(27, 11) = 55.9
myArray(27, 12) = 126.2
myArray(27, 13) = 18.3

myArray(28, 1) = 2022
myArray(28, 2) = 8.6
myArray(28, 3) = 4.7
myArray(28, 4) = 72.1
myArray(28, 5) = 52.2
myArray(28, 6) = 2.9
myArray(28, 7) = 352.4
myArray(28, 8) = 178.4
myArray(28, 9) = 468.7
myArray(28, 10) = 165.9
myArray(28, 11) = 160
myArray(28, 12) = 72.9
myArray(28, 13) = 31.9

myArray(29, 1) = 2023
myArray(29, 2) = 30.5
myArray(29, 3) = 0.1
myArray(29, 4) = 6.4
myArray(29, 5) = 54.6
myArray(29, 6) = 132.9
myArray(29, 7) = 138.1
myArray(29, 8) = 507
myArray(29, 9) = 225
myArray(29, 10) = 166.1
myArray(29, 11) = 39.6
myArray(29, 12) = 122.9
myArray(29, 13) = 106.5

myArray(30, 1) = 2024
myArray(30, 2) = 34.4
myArray(30, 3) = 86.4
myArray(30, 4) = 26.1
myArray(30, 5) = 45
myArray(30, 6) = 157.4
myArray(30, 7) = 120.4
myArray(30, 8) = 556.1
myArray(30, 9) = 197.6
myArray(30, 10) = 354.2
myArray(30, 11) = 150.3
myArray(30, 12) = 62.1
myArray(30, 13) = 16.1


    data_SEOSAN = myArray

End Function

Function data_BUYEO() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 22.6
myArray(1, 3) = 23.5
myArray(1, 4) = 24.4
myArray(1, 5) = 62
myArray(1, 6) = 59.5
myArray(1, 7) = 34.5
myArray(1, 8) = 171.5
myArray(1, 9) = 839
myArray(1, 10) = 46.5
myArray(1, 11) = 22
myArray(1, 12) = 15
myArray(1, 13) = 5.7

myArray(2, 1) = 1996
myArray(2, 2) = 26.4
myArray(2, 3) = 2.8
myArray(2, 4) = 131
myArray(2, 5) = 45
myArray(2, 6) = 33
myArray(2, 7) = 289
myArray(2, 8) = 235
myArray(2, 9) = 67
myArray(2, 10) = 16
myArray(2, 11) = 90.5
myArray(2, 12) = 76
myArray(2, 13) = 35

myArray(3, 1) = 1997
myArray(3, 2) = 9
myArray(3, 3) = 54.9
myArray(3, 4) = 44
myArray(3, 5) = 70
myArray(3, 6) = 229.5
myArray(3, 7) = 236.5
myArray(3, 8) = 404.5
myArray(3, 9) = 263
myArray(3, 10) = 24.5
myArray(3, 11) = 8
myArray(3, 12) = 219.5
myArray(3, 13) = 39.5

myArray(4, 1) = 1998
myArray(4, 2) = 40.6
myArray(4, 3) = 47
myArray(4, 4) = 45
myArray(4, 5) = 200.5
myArray(4, 6) = 130.5
myArray(4, 7) = 324
myArray(4, 8) = 323
myArray(4, 9) = 451.3
myArray(4, 10) = 313.1
myArray(4, 11) = 75.5
myArray(4, 12) = 46.3
myArray(4, 13) = 3.5

myArray(5, 1) = 1999
myArray(5, 2) = 3.5
myArray(5, 3) = 10
myArray(5, 4) = 75.7
myArray(5, 5) = 92.5
myArray(5, 6) = 127.5
myArray(5, 7) = 203
myArray(5, 8) = 149
myArray(5, 9) = 119.5
myArray(5, 10) = 426
myArray(5, 11) = 290
myArray(5, 12) = 15.5
myArray(5, 13) = 17.4

myArray(6, 1) = 2000
myArray(6, 2) = 41.4
myArray(6, 3) = 2.3
myArray(6, 4) = 14.1
myArray(6, 5) = 62
myArray(6, 6) = 40
myArray(6, 7) = 248.5
myArray(6, 8) = 248.5
myArray(6, 9) = 543
myArray(6, 10) = 238.5
myArray(6, 11) = 39
myArray(6, 12) = 29.5
myArray(6, 13) = 13.8

myArray(7, 1) = 2001
myArray(7, 2) = 65
myArray(7, 3) = 69.5
myArray(7, 4) = 9.8
myArray(7, 5) = 25
myArray(7, 6) = 23.5
myArray(7, 7) = 132
myArray(7, 8) = 216
myArray(7, 9) = 98
myArray(7, 10) = 10.5
myArray(7, 11) = 76.5
myArray(7, 12) = 10.5
myArray(7, 13) = 16.3

myArray(8, 1) = 2002
myArray(8, 2) = 72.3
myArray(8, 3) = 6
myArray(8, 4) = 32.5
myArray(8, 5) = 142.5
myArray(8, 6) = 159
myArray(8, 7) = 70.5
myArray(8, 8) = 208
myArray(8, 9) = 358.5
myArray(8, 10) = 57.5
myArray(8, 11) = 78.5
myArray(8, 12) = 31.5
myArray(8, 13) = 57.2

myArray(9, 1) = 2003
myArray(9, 2) = 24.2
myArray(9, 3) = 59
myArray(9, 4) = 52
myArray(9, 5) = 208.5
myArray(9, 6) = 144.5
myArray(9, 7) = 228
myArray(9, 8) = 626.5
myArray(9, 9) = 202
myArray(9, 10) = 167.5
myArray(9, 11) = 24.5
myArray(9, 12) = 29.5
myArray(9, 13) = 13.8

myArray(10, 1) = 2004
myArray(10, 2) = 18.1
myArray(10, 3) = 26.2
myArray(10, 4) = 63.1
myArray(10, 5) = 73.5
myArray(10, 6) = 109
myArray(10, 7) = 388
myArray(10, 8) = 296
myArray(10, 9) = 249
myArray(10, 10) = 176.5
myArray(10, 11) = 1
myArray(10, 12) = 50.5
myArray(10, 13) = 43

myArray(11, 1) = 2005
myArray(11, 2) = 6
myArray(11, 3) = 39
myArray(11, 4) = 26.5
myArray(11, 5) = 75
myArray(11, 6) = 65.5
myArray(11, 7) = 186
myArray(11, 8) = 448.5
myArray(11, 9) = 381.5
myArray(11, 10) = 225.5
myArray(11, 11) = 30.5
myArray(11, 12) = 21
myArray(11, 13) = 22

myArray(12, 1) = 2006
myArray(12, 2) = 30.2
myArray(12, 3) = 29.5
myArray(12, 4) = 7.8
myArray(12, 5) = 99
myArray(12, 6) = 81.5
myArray(12, 7) = 111
myArray(12, 8) = 503
myArray(12, 9) = 83.5
myArray(12, 10) = 37.5
myArray(12, 11) = 15
myArray(12, 12) = 51
myArray(12, 13) = 27.5

myArray(13, 1) = 2007
myArray(13, 2) = 21.8
myArray(13, 3) = 47.8
myArray(13, 4) = 159
myArray(13, 5) = 28
myArray(13, 6) = 104
myArray(13, 7) = 101
myArray(13, 8) = 286
myArray(13, 9) = 319.5
myArray(13, 10) = 502.5
myArray(13, 11) = 37
myArray(13, 12) = 13
myArray(13, 13) = 31.7

myArray(14, 1) = 2008
myArray(14, 2) = 39.6
myArray(14, 3) = 11.2
myArray(14, 4) = 42.2
myArray(14, 5) = 38.8
myArray(14, 6) = 51.6
myArray(14, 7) = 260
myArray(14, 8) = 194.3
myArray(14, 9) = 154
myArray(14, 10) = 48.8
myArray(14, 11) = 24.1
myArray(14, 12) = 14.1
myArray(14, 13) = 23.4

myArray(15, 1) = 2009
myArray(15, 2) = 10.6
myArray(15, 3) = 23.6
myArray(15, 4) = 63.9
myArray(15, 5) = 51
myArray(15, 6) = 135.5
myArray(15, 7) = 113.2
myArray(15, 8) = 408
myArray(15, 9) = 140.2
myArray(15, 10) = 30.5
myArray(15, 11) = 23.7
myArray(15, 12) = 54.5
myArray(15, 13) = 34.9

myArray(16, 1) = 2010
myArray(16, 2) = 37.1
myArray(16, 3) = 89.5
myArray(16, 4) = 94.9
myArray(16, 5) = 69.6
myArray(16, 6) = 140.7
myArray(16, 7) = 36.1
myArray(16, 8) = 262.1
myArray(16, 9) = 431.1
myArray(16, 10) = 149.8
myArray(16, 11) = 17.8
myArray(16, 12) = 18.6
myArray(16, 13) = 31

myArray(17, 1) = 2011
myArray(17, 2) = 3.7
myArray(17, 3) = 60.7
myArray(17, 4) = 16
myArray(17, 5) = 70
myArray(17, 6) = 111.2
myArray(17, 7) = 316
myArray(17, 8) = 599.6
myArray(17, 9) = 618.1
myArray(17, 10) = 104.2
myArray(17, 11) = 26.6
myArray(17, 12) = 81.6
myArray(17, 13) = 7

myArray(18, 1) = 2012
myArray(18, 2) = 16
myArray(18, 3) = 3.2
myArray(18, 4) = 60.2
myArray(18, 5) = 109.3
myArray(18, 6) = 19.5
myArray(18, 7) = 71.3
myArray(18, 8) = 302.9
myArray(18, 9) = 573.3
myArray(18, 10) = 186.2
myArray(18, 11) = 83
myArray(18, 12) = 60.7
myArray(18, 13) = 60.2

myArray(19, 1) = 2013
myArray(19, 2) = 45.4
myArray(19, 3) = 58.7
myArray(19, 4) = 50.3
myArray(19, 5) = 93.7
myArray(19, 6) = 159
myArray(19, 7) = 151.7
myArray(19, 8) = 240.4
myArray(19, 9) = 119.5
myArray(19, 10) = 184.8
myArray(19, 11) = 17.5
myArray(19, 12) = 79.4
myArray(19, 13) = 35.9

myArray(20, 1) = 2014
myArray(20, 2) = 2.2
myArray(20, 3) = 15.3
myArray(20, 4) = 69.3
myArray(20, 5) = 94.1
myArray(20, 6) = 61.5
myArray(20, 7) = 77.8
myArray(20, 8) = 174.7
myArray(20, 9) = 225.1
myArray(20, 10) = 157.5
myArray(20, 11) = 170.5
myArray(20, 12) = 42.4
myArray(20, 13) = 51.7

myArray(21, 1) = 2015
myArray(21, 2) = 35.4
myArray(21, 3) = 35.6
myArray(21, 4) = 42.4
myArray(21, 5) = 99.5
myArray(21, 6) = 53.5
myArray(21, 7) = 92.7
myArray(21, 8) = 119.9
myArray(21, 9) = 56.9
myArray(21, 10) = 22
myArray(21, 11) = 104
myArray(21, 12) = 130
myArray(21, 13) = 56.9

myArray(22, 1) = 2016
myArray(22, 2) = 6.6
myArray(22, 3) = 59.6
myArray(22, 4) = 19
myArray(22, 5) = 164.6
myArray(22, 6) = 121.6
myArray(22, 7) = 49.4
myArray(22, 8) = 341.1
myArray(22, 9) = 33.4
myArray(22, 10) = 133.7
myArray(22, 11) = 120.1
myArray(22, 12) = 17.1
myArray(22, 13) = 63.1

myArray(23, 1) = 2017
myArray(23, 2) = 16
myArray(23, 3) = 28.5
myArray(23, 4) = 8.8
myArray(23, 5) = 78.4
myArray(23, 6) = 35.8
myArray(23, 7) = 51.4
myArray(23, 8) = 326.7
myArray(23, 9) = 358.5
myArray(23, 10) = 97.1
myArray(23, 11) = 51.9
myArray(23, 12) = 22.8
myArray(23, 13) = 36.1

myArray(24, 1) = 2018
myArray(24, 2) = 25
myArray(24, 3) = 43.1
myArray(24, 4) = 99.3
myArray(24, 5) = 156.5
myArray(24, 6) = 116.1
myArray(24, 7) = 107.1
myArray(24, 8) = 278.8
myArray(24, 9) = 277
myArray(24, 10) = 98.3
myArray(24, 11) = 159.2
myArray(24, 12) = 66
myArray(24, 13) = 31.5

myArray(25, 1) = 2019
myArray(25, 2) = 0.5
myArray(25, 3) = 37.6
myArray(25, 4) = 35
myArray(25, 5) = 73.7
myArray(25, 6) = 44.3
myArray(25, 7) = 59.9
myArray(25, 8) = 216.7
myArray(25, 9) = 102.1
myArray(25, 10) = 191.9
myArray(25, 11) = 85.6
myArray(25, 12) = 113.5
myArray(25, 13) = 31.2

myArray(26, 1) = 2020
myArray(26, 2) = 79.6
myArray(26, 3) = 92.4
myArray(26, 4) = 19.3
myArray(26, 5) = 17.7
myArray(26, 6) = 108.5
myArray(26, 7) = 188.4
myArray(26, 8) = 492.6
myArray(26, 9) = 367.8
myArray(26, 10) = 208.9
myArray(26, 11) = 4.4
myArray(26, 12) = 41.8
myArray(26, 13) = 3.4

myArray(27, 1) = 2021
myArray(27, 2) = 32.1
myArray(27, 3) = 18.1
myArray(27, 4) = 95.7
myArray(27, 5) = 42.3
myArray(27, 6) = 136.9
myArray(27, 7) = 76.9
myArray(27, 8) = 187.7
myArray(27, 9) = 227.6
myArray(27, 10) = 187.1
myArray(27, 11) = 36.9
myArray(27, 12) = 73.4
myArray(27, 13) = 8.7

myArray(28, 1) = 2022
myArray(28, 2) = 3.5
myArray(28, 3) = 2.5
myArray(28, 4) = 76.1
myArray(28, 5) = 62.6
myArray(28, 6) = 4
myArray(28, 7) = 123.4
myArray(28, 8) = 168.5
myArray(28, 9) = 615.6
myArray(28, 10) = 87
myArray(28, 11) = 103.7
myArray(28, 12) = 36.4
myArray(28, 13) = 17.8

myArray(29, 1) = 2023
myArray(29, 2) = 35.7
myArray(29, 3) = 4.3
myArray(29, 4) = 13.2
myArray(29, 5) = 60.6
myArray(29, 6) = 248.1
myArray(29, 7) = 122.2
myArray(29, 8) = 880.3
myArray(29, 9) = 300.6
myArray(29, 10) = 303
myArray(29, 11) = 16.7
myArray(29, 12) = 58.1
myArray(29, 13) = 122.2

myArray(30, 1) = 2024
myArray(30, 2) = 43.9
myArray(30, 3) = 135.8
myArray(30, 4) = 49.8
myArray(30, 5) = 58.7
myArray(30, 6) = 134
myArray(30, 7) = 91.2
myArray(30, 8) = 470.8
myArray(30, 9) = 40.8
myArray(30, 10) = 162.6
myArray(30, 11) = 84.5
myArray(30, 12) = 46
myArray(30, 13) = 16.1


    data_BUYEO = myArray

End Function


Function data_BOEUN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 17
myArray(1, 3) = 13.4
myArray(1, 4) = 46.4
myArray(1, 5) = 60.5
myArray(1, 6) = 56
myArray(1, 7) = 45.5
myArray(1, 8) = 126.5
myArray(1, 9) = 508
myArray(1, 10) = 31
myArray(1, 11) = 42
myArray(1, 12) = 32.3
myArray(1, 13) = 5.3

myArray(2, 1) = 1996
myArray(2, 2) = 18.4
myArray(2, 3) = 5.3
myArray(2, 4) = 100.7
myArray(2, 5) = 28.5
myArray(2, 6) = 62.5
myArray(2, 7) = 385.5
myArray(2, 8) = 264
myArray(2, 9) = 97.5
myArray(2, 10) = 29
myArray(2, 11) = 80
myArray(2, 12) = 66.8
myArray(2, 13) = 25.8

myArray(3, 1) = 1997
myArray(3, 2) = 14.1
myArray(3, 3) = 55
myArray(3, 4) = 32
myArray(3, 5) = 49
myArray(3, 6) = 238
myArray(3, 7) = 226
myArray(3, 8) = 378.8
myArray(3, 9) = 402.5
myArray(3, 10) = 46
myArray(3, 11) = 9.5
myArray(3, 12) = 162.7
myArray(3, 13) = 50.1

myArray(4, 1) = 1998
myArray(4, 2) = 22.4
myArray(4, 3) = 28
myArray(4, 4) = 23.7
myArray(4, 5) = 173.5
myArray(4, 6) = 103.5
myArray(4, 7) = 256
myArray(4, 8) = 311.5
myArray(4, 9) = 894
myArray(4, 10) = 180.5
myArray(4, 11) = 54.5
myArray(4, 12) = 33
myArray(4, 13) = 4.5

myArray(5, 1) = 1999
myArray(5, 2) = 1.5
myArray(5, 3) = 6.9
myArray(5, 4) = 72.5
myArray(5, 5) = 118
myArray(5, 6) = 108
myArray(5, 7) = 206
myArray(5, 8) = 136.5
myArray(5, 9) = 249.1
myArray(5, 10) = 306.8
myArray(5, 11) = 144
myArray(5, 12) = 15.6
myArray(5, 13) = 14.3

myArray(6, 1) = 2000
myArray(6, 2) = 34.3
myArray(6, 3) = 3.7
myArray(6, 4) = 20.7
myArray(6, 5) = 60
myArray(6, 6) = 44.5
myArray(6, 7) = 244.6
myArray(6, 8) = 384.2
myArray(6, 9) = 348.5
myArray(6, 10) = 229.5
myArray(6, 11) = 25
myArray(6, 12) = 38
myArray(6, 13) = 16.2

myArray(7, 1) = 2001
myArray(7, 2) = 49.4
myArray(7, 3) = 60.5
myArray(7, 4) = 12.7
myArray(7, 5) = 11.5
myArray(7, 6) = 18
myArray(7, 7) = 259.5
myArray(7, 8) = 139.5
myArray(7, 9) = 127
myArray(7, 10) = 42.5
myArray(7, 11) = 81
myArray(7, 12) = 9
myArray(7, 13) = 23.8

myArray(8, 1) = 2002
myArray(8, 2) = 89.1
myArray(8, 3) = 11
myArray(8, 4) = 35
myArray(8, 5) = 178
myArray(8, 6) = 126
myArray(8, 7) = 49
myArray(8, 8) = 156.5
myArray(8, 9) = 418
myArray(8, 10) = 114.5
myArray(8, 11) = 42
myArray(8, 12) = 20.1
myArray(8, 13) = 45.7

myArray(9, 1) = 2003
myArray(9, 2) = 22.7
myArray(9, 3) = 67
myArray(9, 4) = 44.5
myArray(9, 5) = 200.5
myArray(9, 6) = 159
myArray(9, 7) = 192
myArray(9, 8) = 689
myArray(9, 9) = 310
myArray(9, 10) = 306.5
myArray(9, 11) = 32
myArray(9, 12) = 36.5
myArray(9, 13) = 19.5

myArray(10, 1) = 2004
myArray(10, 2) = 15.4
myArray(10, 3) = 33.8
myArray(10, 4) = 59.9
myArray(10, 5) = 84.2
myArray(10, 6) = 117
myArray(10, 7) = 327
myArray(10, 8) = 295.5
myArray(10, 9) = 203.5
myArray(10, 10) = 136
myArray(10, 11) = 5.5
myArray(10, 12) = 38.5
myArray(10, 13) = 49.1

myArray(11, 1) = 2005
myArray(11, 2) = 9.2
myArray(11, 3) = 20
myArray(11, 4) = 38
myArray(11, 5) = 62.5
myArray(11, 6) = 62
myArray(11, 7) = 215
myArray(11, 8) = 400
myArray(11, 9) = 493.5
myArray(11, 10) = 177.5
myArray(11, 11) = 31
myArray(11, 12) = 15
myArray(11, 13) = 12.6

myArray(12, 1) = 2006
myArray(12, 2) = 27
myArray(12, 3) = 37.8
myArray(12, 4) = 11.9
myArray(12, 5) = 92.5
myArray(12, 6) = 107
myArray(12, 7) = 113
myArray(12, 8) = 511.5
myArray(12, 9) = 143
myArray(12, 10) = 27.5
myArray(12, 11) = 25
myArray(12, 12) = 76.5
myArray(12, 13) = 23.5

myArray(13, 1) = 2007
myArray(13, 2) = 11.6
myArray(13, 3) = 42.5
myArray(13, 4) = 119.8
myArray(13, 5) = 40
myArray(13, 6) = 105
myArray(13, 7) = 142
myArray(13, 8) = 282.5
myArray(13, 9) = 366
myArray(13, 10) = 351.5
myArray(13, 11) = 37
myArray(13, 12) = 10.6
myArray(13, 13) = 23.6

myArray(14, 1) = 2008
myArray(14, 2) = 46.9
myArray(14, 3) = 7
myArray(14, 4) = 28.7
myArray(14, 5) = 23.3
myArray(14, 6) = 83.5
myArray(14, 7) = 152.1
myArray(14, 8) = 212.5
myArray(14, 9) = 311.1
myArray(14, 10) = 51.8
myArray(14, 11) = 19.4
myArray(14, 12) = 11.3
myArray(14, 13) = 14.3

myArray(15, 1) = 2009
myArray(15, 2) = 10.5
myArray(15, 3) = 23.9
myArray(15, 4) = 53
myArray(15, 5) = 37
myArray(15, 6) = 147.5
myArray(15, 7) = 137
myArray(15, 8) = 404
myArray(15, 9) = 124
myArray(15, 10) = 62.5
myArray(15, 11) = 27.2
myArray(15, 12) = 48
myArray(15, 13) = 37.6

myArray(16, 1) = 2010
myArray(16, 2) = 32
myArray(16, 3) = 74.1
myArray(16, 4) = 82.9
myArray(16, 5) = 74
myArray(16, 6) = 108
myArray(16, 7) = 20.6
myArray(16, 8) = 199.6
myArray(16, 9) = 357.7
myArray(16, 10) = 249.7
myArray(16, 11) = 26.8
myArray(16, 12) = 9
myArray(16, 13) = 28.5

myArray(17, 1) = 2011
myArray(17, 2) = 2.6
myArray(17, 3) = 38.9
myArray(17, 4) = 19.8
myArray(17, 5) = 94.5
myArray(17, 6) = 153.8
myArray(17, 7) = 412
myArray(17, 8) = 535.1
myArray(17, 9) = 296.7
myArray(17, 10) = 105.8
myArray(17, 11) = 55.5
myArray(17, 12) = 84.5
myArray(17, 13) = 11.5

myArray(18, 1) = 2012
myArray(18, 2) = 16.1
myArray(18, 3) = 1
myArray(18, 4) = 79
myArray(18, 5) = 92.4
myArray(18, 6) = 44.6
myArray(18, 7) = 79.4
myArray(18, 8) = 294.6
myArray(18, 9) = 488.9
myArray(18, 10) = 218.5
myArray(18, 11) = 83.6
myArray(18, 12) = 68.2
myArray(18, 13) = 56

myArray(19, 1) = 2013
myArray(19, 2) = 45.2
myArray(19, 3) = 37.2
myArray(19, 4) = 50.9
myArray(19, 5) = 90.3
myArray(19, 6) = 107.8
myArray(19, 7) = 160.9
myArray(19, 8) = 245.2
myArray(19, 9) = 114
myArray(19, 10) = 139.9
myArray(19, 11) = 32
myArray(19, 12) = 67.2
myArray(19, 13) = 35.3

myArray(20, 1) = 2014
myArray(20, 2) = 8
myArray(20, 3) = 5.5
myArray(20, 4) = 72.2
myArray(20, 5) = 46.7
myArray(20, 6) = 53
myArray(20, 7) = 103.4
myArray(20, 8) = 164.7
myArray(20, 9) = 288.7
myArray(20, 10) = 106.7
myArray(20, 11) = 171.3
myArray(20, 12) = 43.2
myArray(20, 13) = 25.8

myArray(21, 1) = 2015
myArray(21, 2) = 25.1
myArray(21, 3) = 31.8
myArray(21, 4) = 45
myArray(21, 5) = 92.1
myArray(21, 6) = 31.8
myArray(21, 7) = 73.4
myArray(21, 8) = 156.7
myArray(21, 9) = 85.2
myArray(21, 10) = 38.8
myArray(21, 11) = 84.5
myArray(21, 12) = 109.8
myArray(21, 13) = 42.8

myArray(22, 1) = 2016
myArray(22, 2) = 9.7
myArray(22, 3) = 39.8
myArray(22, 4) = 45.5
myArray(22, 5) = 149.2
myArray(22, 6) = 88.5
myArray(22, 7) = 48.8
myArray(22, 8) = 494.9
myArray(22, 9) = 53
myArray(22, 10) = 149.4
myArray(22, 11) = 135.7
myArray(22, 12) = 33.5
myArray(22, 13) = 43.6

myArray(23, 1) = 2017
myArray(23, 2) = 16.4
myArray(23, 3) = 47
myArray(23, 4) = 16.9
myArray(23, 5) = 61
myArray(23, 6) = 28
myArray(23, 7) = 87.6
myArray(23, 8) = 572
myArray(23, 9) = 315.8
myArray(23, 10) = 108.6
myArray(23, 11) = 28.5
myArray(23, 12) = 13.4
myArray(23, 13) = 27.6

myArray(24, 1) = 2018
myArray(24, 2) = 25
myArray(24, 3) = 28.5
myArray(24, 4) = 93.5
myArray(24, 5) = 139.2
myArray(24, 6) = 103.9
myArray(24, 7) = 75.3
myArray(24, 8) = 224.9
myArray(24, 9) = 386
myArray(24, 10) = 134.7
myArray(24, 11) = 108.3
myArray(24, 12) = 52.2
myArray(24, 13) = 38

myArray(25, 1) = 2019
myArray(25, 2) = 0.9
myArray(25, 3) = 39.8
myArray(25, 4) = 29
myArray(25, 5) = 100.8
myArray(25, 6) = 42.8
myArray(25, 7) = 73.9
myArray(25, 8) = 226.1
myArray(25, 9) = 132.7
myArray(25, 10) = 186.2
myArray(25, 11) = 100.6
myArray(25, 12) = 86.2
myArray(25, 13) = 27.2

myArray(26, 1) = 2020
myArray(26, 2) = 71.8
myArray(26, 3) = 80.8
myArray(26, 4) = 23.6
myArray(26, 5) = 35.9
myArray(26, 6) = 89.1
myArray(26, 7) = 171.2
myArray(26, 8) = 500.8
myArray(26, 9) = 587.8
myArray(26, 10) = 162.2
myArray(26, 11) = 3
myArray(26, 12) = 37
myArray(26, 13) = 5.6

myArray(27, 1) = 2021
myArray(27, 2) = 19.8
myArray(27, 3) = 14.1
myArray(27, 4) = 93
myArray(27, 5) = 52.5
myArray(27, 6) = 154.5
myArray(27, 7) = 76.8
myArray(27, 8) = 163.9
myArray(27, 9) = 275.8
myArray(27, 10) = 162.2
myArray(27, 11) = 33.8
myArray(27, 12) = 44.9
myArray(27, 13) = 6.4

myArray(28, 1) = 2022
myArray(28, 2) = 4.3
myArray(28, 3) = 4.3
myArray(28, 4) = 98.2
myArray(28, 5) = 61.1
myArray(28, 6) = 5.6
myArray(28, 7) = 98.3
myArray(28, 8) = 160.5
myArray(28, 9) = 391.9
myArray(28, 10) = 71.7
myArray(28, 11) = 89.7
myArray(28, 12) = 39.9
myArray(28, 13) = 18

myArray(29, 1) = 2023
myArray(29, 2) = 27.3
myArray(29, 3) = 5.3
myArray(29, 4) = 25
myArray(29, 5) = 51.9
myArray(29, 6) = 146.4
myArray(29, 7) = 169.3
myArray(29, 8) = 796.9
myArray(29, 9) = 242.9
myArray(29, 10) = 243.2
myArray(29, 11) = 14.5
myArray(29, 12) = 39.6
myArray(29, 13) = 108.1

myArray(30, 1) = 2024
myArray(30, 2) = 34.1
myArray(30, 3) = 79.6
myArray(30, 4) = 50.7
myArray(30, 5) = 45.8
myArray(30, 6) = 128.4
myArray(30, 7) = 70.1
myArray(30, 8) = 513.8
myArray(30, 9) = 47.6
myArray(30, 10) = 148.1
myArray(30, 11) = 99.4
myArray(30, 12) = 29.6
myArray(30, 13) = 5.6


    data_BOEUN = myArray

End Function

Function data_BORYUNG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 15.7
myArray(1, 3) = 11
myArray(1, 4) = 19.6
myArray(1, 5) = 65.5
myArray(1, 6) = 49.5
myArray(1, 7) = 26.5
myArray(1, 8) = 144.5
myArray(1, 9) = 996.5
myArray(1, 10) = 70.5
myArray(1, 11) = 24.5
myArray(1, 12) = 23
myArray(1, 13) = 12.7

myArray(2, 1) = 1996
myArray(2, 2) = 33.4
myArray(2, 3) = 6.8
myArray(2, 4) = 104.5
myArray(2, 5) = 34
myArray(2, 6) = 22.5
myArray(2, 7) = 235
myArray(2, 8) = 192.5
myArray(2, 9) = 44.5
myArray(2, 10) = 14
myArray(2, 11) = 106.5
myArray(2, 12) = 74.2
myArray(2, 13) = 31.7

myArray(3, 1) = 1997
myArray(3, 2) = 15.1
myArray(3, 3) = 38.4
myArray(3, 4) = 30.5
myArray(3, 5) = 57.5
myArray(3, 6) = 203
myArray(3, 7) = 272
myArray(3, 8) = 353.5
myArray(3, 9) = 211.5
myArray(3, 10) = 23
myArray(3, 11) = 10
myArray(3, 12) = 193.5
myArray(3, 13) = 34.3

myArray(4, 1) = 1998
myArray(4, 2) = 29.9
myArray(4, 3) = 40.2
myArray(4, 4) = 30.5
myArray(4, 5) = 138
myArray(4, 6) = 100
myArray(4, 7) = 209.5
myArray(4, 8) = 263
myArray(4, 9) = 341.7
myArray(4, 10) = 150.3
myArray(4, 11) = 61
myArray(4, 12) = 29.3
myArray(4, 13) = 3.8

myArray(5, 1) = 1999
myArray(5, 2) = 7.9
myArray(5, 3) = 9.5
myArray(5, 4) = 71
myArray(5, 5) = 88.5
myArray(5, 6) = 124.5
myArray(5, 7) = 192.5
myArray(5, 8) = 98
myArray(5, 9) = 180
myArray(5, 10) = 292.5
myArray(5, 11) = 169
myArray(5, 12) = 24.9
myArray(5, 13) = 25.8

myArray(6, 1) = 2000
myArray(6, 2) = 42.1
myArray(6, 3) = 3.2
myArray(6, 4) = 7
myArray(6, 5) = 35
myArray(6, 6) = 53.5
myArray(6, 7) = 159.5
myArray(6, 8) = 155
myArray(6, 9) = 701.5
myArray(6, 10) = 241
myArray(6, 11) = 46
myArray(6, 12) = 39.5
myArray(6, 13) = 32.1

myArray(7, 1) = 2001
myArray(7, 2) = 73.3
myArray(7, 3) = 46
myArray(7, 4) = 15.9
myArray(7, 5) = 26
myArray(7, 6) = 17
myArray(7, 7) = 129
myArray(7, 8) = 286.5
myArray(7, 9) = 170
myArray(7, 10) = 10
myArray(7, 11) = 85
myArray(7, 12) = 13
myArray(7, 13) = 32

myArray(8, 1) = 2002
myArray(8, 2) = 50.8
myArray(8, 3) = 5.5
myArray(8, 4) = 32
myArray(8, 5) = 169
myArray(8, 6) = 155.5
myArray(8, 7) = 72
myArray(8, 8) = 217.5
myArray(8, 9) = 477
myArray(8, 10) = 27
myArray(8, 11) = 134
myArray(8, 12) = 61.1
myArray(8, 13) = 51.8

myArray(9, 1) = 2003
myArray(9, 2) = 30.7
myArray(9, 3) = 44.5
myArray(9, 4) = 39.5
myArray(9, 5) = 168.5
myArray(9, 6) = 78.5
myArray(9, 7) = 153
myArray(9, 8) = 309.5
myArray(9, 9) = 310
myArray(9, 10) = 128
myArray(9, 11) = 23
myArray(9, 12) = 45.5
myArray(9, 13) = 13

myArray(10, 1) = 2004
myArray(10, 2) = 22.1
myArray(10, 3) = 28.5
myArray(10, 4) = 45.7
myArray(10, 5) = 58
myArray(10, 6) = 105.5
myArray(10, 7) = 234.5
myArray(10, 8) = 263.5
myArray(10, 9) = 164
myArray(10, 10) = 195
myArray(10, 11) = 4
myArray(10, 12) = 56.5
myArray(10, 13) = 38.9

myArray(11, 1) = 2005
myArray(11, 2) = 5.8
myArray(11, 3) = 35.8
myArray(11, 4) = 30
myArray(11, 5) = 73.5
myArray(11, 6) = 48.5
myArray(11, 7) = 156
myArray(11, 8) = 260.5
myArray(11, 9) = 291.5
myArray(11, 10) = 282.5
myArray(11, 11) = 21
myArray(11, 12) = 18
myArray(11, 13) = 43.4

myArray(12, 1) = 2006
myArray(12, 2) = 27
myArray(12, 3) = 25.9
myArray(12, 4) = 10.6
myArray(12, 5) = 81.5
myArray(12, 6) = 94.5
myArray(12, 7) = 114.5
myArray(12, 8) = 321
myArray(12, 9) = 21.5
myArray(12, 10) = 23.5
myArray(12, 11) = 24.5
myArray(12, 12) = 61.5
myArray(12, 13) = 25.4

myArray(13, 1) = 2007
myArray(13, 2) = 23.4
myArray(13, 3) = 29.8
myArray(13, 4) = 102
myArray(13, 5) = 29.5
myArray(13, 6) = 79
myArray(13, 7) = 85
myArray(13, 8) = 214
myArray(13, 9) = 239.5
myArray(13, 10) = 384
myArray(13, 11) = 59
myArray(13, 12) = 17.5
myArray(13, 13) = 33.1

myArray(14, 1) = 2008
myArray(14, 2) = 20.9
myArray(14, 3) = 10.8
myArray(14, 4) = 48.2
myArray(14, 5) = 40.5
myArray(14, 6) = 78.9
myArray(14, 7) = 101.3
myArray(14, 8) = 257.2
myArray(14, 9) = 119.5
myArray(14, 10) = 46.9
myArray(14, 11) = 26.7
myArray(14, 12) = 37.6
myArray(14, 13) = 25

myArray(15, 1) = 2009
myArray(15, 2) = 18.5
myArray(15, 3) = 23.3
myArray(15, 4) = 55.1
myArray(15, 5) = 41.5
myArray(15, 6) = 154.5
myArray(15, 7) = 115.1
myArray(15, 8) = 320.9
myArray(15, 9) = 176.6
myArray(15, 10) = 25.5
myArray(15, 11) = 39.5
myArray(15, 12) = 52.9
myArray(15, 13) = 58

myArray(16, 1) = 2010
myArray(16, 2) = 30.1
myArray(16, 3) = 73.5
myArray(16, 4) = 75.9
myArray(16, 5) = 58.5
myArray(16, 6) = 122.8
myArray(16, 7) = 60.8
myArray(16, 8) = 396.5
myArray(16, 9) = 402.7
myArray(16, 10) = 213.1
myArray(16, 11) = 19.2
myArray(16, 12) = 16.3
myArray(16, 13) = 32.9

myArray(17, 1) = 2011
myArray(17, 2) = 11.1
myArray(17, 3) = 37.5
myArray(17, 4) = 18
myArray(17, 5) = 72.1
myArray(17, 6) = 115.3
myArray(17, 7) = 318
myArray(17, 8) = 723.1
myArray(17, 9) = 289.4
myArray(17, 10) = 70.8
myArray(17, 11) = 13.9
myArray(17, 12) = 61.3
myArray(17, 13) = 12.5

myArray(18, 1) = 2012
myArray(18, 2) = 24.2
myArray(18, 3) = 9.2
myArray(18, 4) = 45
myArray(18, 5) = 68.9
myArray(18, 6) = 14.6
myArray(18, 7) = 76.8
myArray(18, 8) = 231.1
myArray(18, 9) = 450.2
myArray(18, 10) = 207.7
myArray(18, 11) = 65
myArray(18, 12) = 61.1
myArray(18, 13) = 65.2

myArray(19, 1) = 2013
myArray(19, 2) = 28.4
myArray(19, 3) = 40.7
myArray(19, 4) = 53.4
myArray(19, 5) = 68.2
myArray(19, 6) = 116.6
myArray(19, 7) = 159.9
myArray(19, 8) = 267.5
myArray(19, 9) = 214.6
myArray(19, 10) = 320
myArray(19, 11) = 10.9
myArray(19, 12) = 81.1
myArray(19, 13) = 26.4

myArray(20, 1) = 2014
myArray(20, 2) = 3.4
myArray(20, 3) = 20.5
myArray(20, 4) = 56.3
myArray(20, 5) = 70
myArray(20, 6) = 47.1
myArray(20, 7) = 125.8
myArray(20, 8) = 104
myArray(20, 9) = 168.5
myArray(20, 10) = 152
myArray(20, 11) = 156
myArray(20, 12) = 39.9
myArray(20, 13) = 66.6

myArray(21, 1) = 2015
myArray(21, 2) = 29.9
myArray(21, 3) = 23.4
myArray(21, 4) = 30.9
myArray(21, 5) = 129.7
myArray(21, 6) = 38.8
myArray(21, 7) = 83.9
myArray(21, 8) = 94.7
myArray(21, 9) = 30.2
myArray(21, 10) = 13.3
myArray(21, 11) = 90
myArray(21, 12) = 155.6
myArray(21, 13) = 65

myArray(22, 1) = 2016
myArray(22, 2) = 7.8
myArray(22, 3) = 54.2
myArray(22, 4) = 18.7
myArray(22, 5) = 105.1
myArray(22, 6) = 146.5
myArray(22, 7) = 23.7
myArray(22, 8) = 200.2
myArray(22, 9) = 5.1
myArray(22, 10) = 73.4
myArray(22, 11) = 108
myArray(22, 12) = 5.6
myArray(22, 13) = 44.5

myArray(23, 1) = 2017
myArray(23, 2) = 14.8
myArray(23, 3) = 30.2
myArray(23, 4) = 14.4
myArray(23, 5) = 57.6
myArray(23, 6) = 58.9
myArray(23, 7) = 21.1
myArray(23, 8) = 278.1
myArray(23, 9) = 210
myArray(23, 10) = 90
myArray(23, 11) = 26.6
myArray(23, 12) = 15.9
myArray(23, 13) = 38.6

myArray(24, 1) = 2018
myArray(24, 2) = 15
myArray(24, 3) = 33.6
myArray(24, 4) = 92
myArray(24, 5) = 128.1
myArray(24, 6) = 104.5
myArray(24, 7) = 71
myArray(24, 8) = 262.7
myArray(24, 9) = 239.6
myArray(24, 10) = 158.2
myArray(24, 11) = 154.7
myArray(24, 12) = 46.7
myArray(24, 13) = 31.1

myArray(25, 1) = 2019
myArray(25, 2) = 1.9
myArray(25, 3) = 17.8
myArray(25, 4) = 18.2
myArray(25, 5) = 71.9
myArray(25, 6) = 31.3
myArray(25, 7) = 56
myArray(25, 8) = 149
myArray(25, 9) = 131.3
myArray(25, 10) = 118.7
myArray(25, 11) = 63.9
myArray(25, 12) = 130.6
myArray(25, 13) = 31.3

myArray(26, 1) = 2020
myArray(26, 2) = 49.4
myArray(26, 3) = 75.3
myArray(26, 4) = 22.8
myArray(26, 5) = 16.5
myArray(26, 6) = 92.4
myArray(26, 7) = 139.7
myArray(26, 8) = 345.9
myArray(26, 9) = 321.5
myArray(26, 10) = 177.1
myArray(26, 11) = 16.2
myArray(26, 12) = 35.4
myArray(26, 13) = 9.7

myArray(27, 1) = 2021
myArray(27, 2) = 32
myArray(27, 3) = 18.7
myArray(27, 4) = 76.1
myArray(27, 5) = 43.4
myArray(27, 6) = 110
myArray(27, 7) = 55
myArray(27, 8) = 131.3
myArray(27, 9) = 253.7
myArray(27, 10) = 215.9
myArray(27, 11) = 39.6
myArray(27, 12) = 117.8
myArray(27, 13) = 14.4

myArray(28, 1) = 2022
myArray(28, 2) = 8.4
myArray(28, 3) = 5.3
myArray(28, 4) = 60.9
myArray(28, 5) = 34.8
myArray(28, 6) = 5.7
myArray(28, 7) = 225
myArray(28, 8) = 119.7
myArray(28, 9) = 637.1
myArray(28, 10) = 102
myArray(28, 11) = 112
myArray(28, 12) = 23.3
myArray(28, 13) = 14

myArray(29, 1) = 2023
myArray(29, 2) = 19.2
myArray(29, 3) = 0.4
myArray(29, 4) = 7.2
myArray(29, 5) = 42.4
myArray(29, 6) = 190.8
myArray(29, 7) = 95.1
myArray(29, 8) = 772.2
myArray(29, 9) = 107.8
myArray(29, 10) = 288.4
myArray(29, 11) = 25.5
myArray(29, 12) = 65.2
myArray(29, 13) = 112.5

myArray(30, 1) = 2024
myArray(30, 2) = 29.5
myArray(30, 3) = 92.6
myArray(30, 4) = 35.8
myArray(30, 5) = 48.8
myArray(30, 6) = 130.9
myArray(30, 7) = 89
myArray(30, 8) = 557.3
myArray(30, 9) = 149.7
myArray(30, 10) = 234.9
myArray(30, 11) = 58.3
myArray(30, 12) = 49
myArray(30, 13) = 18.5


    data_BORYUNG = myArray

End Function

Function data_DAEJEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 23.5
myArray(1, 3) = 16.9
myArray(1, 4) = 33.8
myArray(1, 5) = 54.7
myArray(1, 6) = 62.2
myArray(1, 7) = 33.6
myArray(1, 8) = 155.4
myArray(1, 9) = 641.9
myArray(1, 10) = 53.4
myArray(1, 11) = 36
myArray(1, 12) = 17.5
myArray(1, 13) = 7.3

myArray(2, 1) = 1996
myArray(2, 2) = 32.7
myArray(2, 3) = 4.4
myArray(2, 4) = 138
myArray(2, 5) = 49.8
myArray(2, 6) = 62.9
myArray(2, 7) = 411.4
myArray(2, 8) = 257.7
myArray(2, 9) = 114.4
myArray(2, 10) = 11.4
myArray(2, 11) = 90.8
myArray(2, 12) = 77.1
myArray(2, 13) = 28.6

myArray(3, 1) = 1997
myArray(3, 2) = 15.6
myArray(3, 3) = 51.1
myArray(3, 4) = 37.1
myArray(3, 5) = 55.4
myArray(3, 6) = 200.9
myArray(3, 7) = 267.5
myArray(3, 8) = 424.2
myArray(3, 9) = 463.5
myArray(3, 10) = 30.2
myArray(3, 11) = 7.7
myArray(3, 12) = 168.2
myArray(3, 13) = 44.5

myArray(4, 1) = 1998
myArray(4, 2) = 33.3
myArray(4, 3) = 36.3
myArray(4, 4) = 31.1
myArray(4, 5) = 154.3
myArray(4, 6) = 119.5
myArray(4, 7) = 297.2
myArray(4, 8) = 256.1
myArray(4, 9) = 781.7
myArray(4, 10) = 254.7
myArray(4, 11) = 71.5
myArray(4, 12) = 31.6
myArray(4, 13) = 2.7

myArray(5, 1) = 1999
myArray(5, 2) = 1.8
myArray(5, 3) = 12.2
myArray(5, 4) = 79.4
myArray(5, 5) = 103
myArray(5, 6) = 116.8
myArray(5, 7) = 245.7
myArray(5, 8) = 137.8
myArray(5, 9) = 203
myArray(5, 10) = 359.5
myArray(5, 11) = 171.6
myArray(5, 12) = 16.5
myArray(5, 13) = 7.9

myArray(6, 1) = 2000
myArray(6, 2) = 27.5
myArray(6, 3) = 4.1
myArray(6, 4) = 17.8
myArray(6, 5) = 67.8
myArray(6, 6) = 54.3
myArray(6, 7) = 238.3
myArray(6, 8) = 470.1
myArray(6, 9) = 473.6
myArray(6, 10) = 263.2
myArray(6, 11) = 24.6
myArray(6, 12) = 44.6
myArray(6, 13) = 21.6

myArray(7, 1) = 2001
myArray(7, 2) = 61.2
myArray(7, 3) = 70
myArray(7, 4) = 16
myArray(7, 5) = 20.4
myArray(7, 6) = 30.2
myArray(7, 7) = 234.2
myArray(7, 8) = 171
myArray(7, 9) = 78.1
myArray(7, 10) = 25.2
myArray(7, 11) = 91.2
myArray(7, 12) = 10.8
myArray(7, 13) = 20.4

myArray(8, 1) = 2002
myArray(8, 2) = 92.1
myArray(8, 3) = 12
myArray(8, 4) = 33.5
myArray(8, 5) = 155.5
myArray(8, 6) = 130.5
myArray(8, 7) = 55.4
myArray(8, 8) = 149.1
myArray(8, 9) = 538.8
myArray(8, 10) = 77
myArray(8, 11) = 67.8
myArray(8, 12) = 24
myArray(8, 13) = 43

myArray(9, 1) = 2003
myArray(9, 2) = 11.2
myArray(9, 3) = 59.2
myArray(9, 4) = 44.2
myArray(9, 5) = 217.5
myArray(9, 6) = 119.5
myArray(9, 7) = 186.4
myArray(9, 8) = 576.3
myArray(9, 9) = 254.9
myArray(9, 10) = 208.5
myArray(9, 11) = 21.5
myArray(9, 12) = 32.6
myArray(9, 13) = 17.1

myArray(10, 1) = 2004
myArray(10, 2) = 10.9
myArray(10, 3) = 30.6
myArray(10, 4) = 83.2
myArray(10, 5) = 73.1
myArray(10, 6) = 109
myArray(10, 7) = 383.5
myArray(10, 8) = 391
myArray(10, 9) = 198.3
myArray(10, 10) = 133.7
myArray(10, 11) = 5
myArray(10, 12) = 37.1
myArray(10, 13) = 41.1

myArray(11, 1) = 2005
myArray(11, 2) = 6
myArray(11, 3) = 37.5
myArray(11, 4) = 38.8
myArray(11, 5) = 48.5
myArray(11, 6) = 60.5
myArray(11, 7) = 209.6
myArray(11, 8) = 463.3
myArray(11, 9) = 499.5
myArray(11, 10) = 226.4
myArray(11, 11) = 30.5
myArray(11, 12) = 20.3
myArray(11, 13) = 15.2

myArray(12, 1) = 2006
myArray(12, 2) = 31.2
myArray(12, 3) = 33.1
myArray(12, 4) = 8.1
myArray(12, 5) = 94.2
myArray(12, 6) = 119.7
myArray(12, 7) = 131
myArray(12, 8) = 531
myArray(12, 9) = 113.6
myArray(12, 10) = 24.1
myArray(12, 11) = 19.3
myArray(12, 12) = 60
myArray(12, 13) = 29.9

myArray(13, 1) = 2007
myArray(13, 2) = 14
myArray(13, 3) = 45
myArray(13, 4) = 117.5
myArray(13, 5) = 28.6
myArray(13, 6) = 130.1
myArray(13, 7) = 133
myArray(13, 8) = 275.7
myArray(13, 9) = 373
myArray(13, 10) = 549.9
myArray(13, 11) = 47.4
myArray(13, 12) = 9.8
myArray(13, 13) = 26.9

myArray(14, 1) = 2008
myArray(14, 2) = 45.3
myArray(14, 3) = 9.1
myArray(14, 4) = 29.1
myArray(14, 5) = 34.4
myArray(14, 6) = 59.2
myArray(14, 7) = 148.3
myArray(14, 8) = 253.4
myArray(14, 9) = 325.2
myArray(14, 10) = 85.5
myArray(14, 11) = 19.6
myArray(14, 12) = 12.1
myArray(14, 13) = 16.4

myArray(15, 1) = 2009
myArray(15, 2) = 15.4
myArray(15, 3) = 27.5
myArray(15, 4) = 60.3
myArray(15, 5) = 34.5
myArray(15, 6) = 124.3
myArray(15, 7) = 87.3
myArray(15, 8) = 429.2
myArray(15, 9) = 148.3
myArray(15, 10) = 46
myArray(15, 11) = 24.7
myArray(15, 12) = 54.7
myArray(15, 13) = 38.2

myArray(16, 1) = 2010
myArray(16, 2) = 46.4
myArray(16, 3) = 81.5
myArray(16, 4) = 100.1
myArray(16, 5) = 88.5
myArray(16, 6) = 117.6
myArray(16, 7) = 65.6
myArray(16, 8) = 223.1
myArray(16, 9) = 376.4
myArray(16, 10) = 250.5
myArray(16, 11) = 17.9
myArray(16, 12) = 16.4
myArray(16, 13) = 35.7

myArray(17, 1) = 2011
myArray(17, 2) = 4
myArray(17, 3) = 44.8
myArray(17, 4) = 19
myArray(17, 5) = 71
myArray(17, 6) = 162
myArray(17, 7) = 391.6
myArray(17, 8) = 587.3
myArray(17, 9) = 420.3
myArray(17, 10) = 91.7
myArray(17, 11) = 37
myArray(17, 12) = 103.2
myArray(17, 13) = 11.5

myArray(18, 1) = 2012
myArray(18, 2) = 16.4
myArray(18, 3) = 2.5
myArray(18, 4) = 54.6
myArray(18, 5) = 66.2
myArray(18, 6) = 24
myArray(18, 7) = 57.8
myArray(18, 8) = 277.6
myArray(18, 9) = 463.6
myArray(18, 10) = 242.5
myArray(18, 11) = 81.3
myArray(18, 12) = 58.4
myArray(18, 13) = 64.6

myArray(19, 1) = 2013
myArray(19, 2) = 46.2
myArray(19, 3) = 54.2
myArray(19, 4) = 52.8
myArray(19, 5) = 86.8
myArray(19, 6) = 110.4
myArray(19, 7) = 162.6
myArray(19, 8) = 218.7
myArray(19, 9) = 126.6
myArray(19, 10) = 146.4
myArray(19, 11) = 19.6
myArray(19, 12) = 63.1
myArray(19, 13) = 32.8

myArray(20, 1) = 2014
myArray(20, 2) = 6.5
myArray(20, 3) = 8.5
myArray(20, 4) = 67.2
myArray(20, 5) = 59.4
myArray(20, 6) = 49.7
myArray(20, 7) = 143.7
myArray(20, 8) = 177.2
myArray(20, 9) = 240.9
myArray(20, 10) = 118
myArray(20, 11) = 169.4
myArray(20, 12) = 40.7
myArray(20, 13) = 36.5

myArray(21, 1) = 2015
myArray(21, 2) = 31.5
myArray(21, 3) = 27
myArray(21, 4) = 44.7
myArray(21, 5) = 95.2
myArray(21, 6) = 28.9
myArray(21, 7) = 119.8
myArray(21, 8) = 145.6
myArray(21, 9) = 51.6
myArray(21, 10) = 18.5
myArray(21, 11) = 94.1
myArray(21, 12) = 126.1
myArray(21, 13) = 39.7

myArray(22, 1) = 2016
myArray(22, 2) = 11.6
myArray(22, 3) = 45.8
myArray(22, 4) = 40.3
myArray(22, 5) = 154.9
myArray(22, 6) = 85.1
myArray(22, 7) = 62.5
myArray(22, 8) = 367.9
myArray(22, 9) = 57.4
myArray(22, 10) = 196
myArray(22, 11) = 122.6
myArray(22, 12) = 29.5
myArray(22, 13) = 54.8

myArray(23, 1) = 2017
myArray(23, 2) = 15
myArray(23, 3) = 42
myArray(23, 4) = 11.6
myArray(23, 5) = 77.7
myArray(23, 6) = 29.3
myArray(23, 7) = 35.3
myArray(23, 8) = 434.5
myArray(23, 9) = 293.8
myArray(23, 10) = 111.4
myArray(23, 11) = 28.3
myArray(23, 12) = 15.1
myArray(23, 13) = 33.5

myArray(24, 1) = 2018
myArray(24, 2) = 23.9
myArray(24, 3) = 40.5
myArray(24, 4) = 108.4
myArray(24, 5) = 155.3
myArray(24, 6) = 95.9
myArray(24, 7) = 115.8
myArray(24, 8) = 226.9
myArray(24, 9) = 408.6
myArray(24, 10) = 149.4
myArray(24, 11) = 133.9
myArray(24, 12) = 49.8
myArray(24, 13) = 33.7

myArray(25, 1) = 2019
myArray(25, 2) = 1.7
myArray(25, 3) = 46.3
myArray(25, 4) = 33.7
myArray(25, 5) = 91.6
myArray(25, 6) = 35.6
myArray(25, 7) = 77.9
myArray(25, 8) = 199
myArray(25, 9) = 104.3
myArray(25, 10) = 167
myArray(25, 11) = 106.1
myArray(25, 12) = 94
myArray(25, 13) = 27

myArray(26, 1) = 2020
myArray(26, 2) = 78.5
myArray(26, 3) = 91.2
myArray(26, 4) = 24.4
myArray(26, 5) = 17.8
myArray(26, 6) = 80.4
myArray(26, 7) = 192.5
myArray(26, 8) = 544.9
myArray(26, 9) = 361.6
myArray(26, 10) = 173.6
myArray(26, 11) = 3.2
myArray(26, 12) = 41.8
myArray(26, 13) = 4.1

myArray(27, 1) = 2021
myArray(27, 2) = 23.6
myArray(27, 3) = 14.1
myArray(27, 4) = 95.2
myArray(27, 5) = 47.4
myArray(27, 6) = 134.2
myArray(27, 7) = 105.9
myArray(27, 8) = 151.8
myArray(27, 9) = 289.2
myArray(27, 10) = 161.2
myArray(27, 11) = 40.8
myArray(27, 12) = 41.7
myArray(27, 13) = 4.4

myArray(28, 1) = 2022
myArray(28, 2) = 1.2
myArray(28, 3) = 1.4
myArray(28, 4) = 74
myArray(28, 5) = 69.7
myArray(28, 6) = 8.1
myArray(28, 7) = 117.6
myArray(28, 8) = 195
myArray(28, 9) = 496.1
myArray(28, 10) = 90.2
myArray(28, 11) = 89.3
myArray(28, 12) = 45.8
myArray(28, 13) = 14.7

myArray(29, 1) = 2023
myArray(29, 2) = 28.4
myArray(29, 3) = 5.4
myArray(29, 4) = 23.8
myArray(29, 5) = 54.5
myArray(29, 6) = 192.9
myArray(29, 7) = 147.5
myArray(29, 8) = 776.3
myArray(29, 9) = 326.9
myArray(29, 10) = 310.2
myArray(29, 11) = 12.2
myArray(29, 12) = 40.3
myArray(29, 13) = 124.1

myArray(30, 1) = 2024
myArray(30, 2) = 47.8
myArray(30, 3) = 93.9
myArray(30, 4) = 57.9
myArray(30, 5) = 32.7
myArray(30, 6) = 126.8
myArray(30, 7) = 76.9
myArray(30, 8) = 485.1
myArray(30, 9) = 87.3
myArray(30, 10) = 204.6
myArray(30, 11) = 109.2
myArray(30, 12) = 34.6
myArray(30, 13) = 3.7


    data_DAEJEON = myArray

End Function



Function data_GEUMSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 23.2
myArray(1, 3) = 17.1
myArray(1, 4) = 46.9
myArray(1, 5) = 65.5
myArray(1, 6) = 35.5
myArray(1, 7) = 54
myArray(1, 8) = 83.5
myArray(1, 9) = 579.5
myArray(1, 10) = 47.5
myArray(1, 11) = 23.5
myArray(1, 12) = 31
myArray(1, 13) = 4.6

myArray(2, 1) = 1996
myArray(2, 2) = 25.4
myArray(2, 3) = 2.9
myArray(2, 4) = 123
myArray(2, 5) = 42.5
myArray(2, 6) = 37.5
myArray(2, 7) = 546
myArray(2, 8) = 174
myArray(2, 9) = 130
myArray(2, 10) = 12.5
myArray(2, 11) = 75.5
myArray(2, 12) = 89.8
myArray(2, 13) = 43

myArray(3, 1) = 1997
myArray(3, 2) = 21.3
myArray(3, 3) = 48.2
myArray(3, 4) = 34
myArray(3, 5) = 58
myArray(3, 6) = 170.5
myArray(3, 7) = 238.5
myArray(3, 8) = 444.5
myArray(3, 9) = 246.5
myArray(3, 10) = 89
myArray(3, 11) = 9
myArray(3, 12) = 160
myArray(3, 13) = 49

myArray(4, 1) = 1998
myArray(4, 2) = 38.4
myArray(4, 3) = 53.9
myArray(4, 4) = 25.6
myArray(4, 5) = 177.5
myArray(4, 6) = 98.5
myArray(4, 7) = 278.5
myArray(4, 8) = 184
myArray(4, 9) = 520
myArray(4, 10) = 237.3
myArray(4, 11) = 49
myArray(4, 12) = 46.1
myArray(4, 13) = 6.8

myArray(5, 1) = 1999
myArray(5, 2) = 5.3
myArray(5, 3) = 22.9
myArray(5, 4) = 73
myArray(5, 5) = 91.5
myArray(5, 6) = 117.5
myArray(5, 7) = 198
myArray(5, 8) = 114.5
myArray(5, 9) = 167.5
myArray(5, 10) = 289.5
myArray(5, 11) = 125
myArray(5, 12) = 16.4
myArray(5, 13) = 10.3

myArray(6, 1) = 2000
myArray(6, 2) = 36.2
myArray(6, 3) = 2.9
myArray(6, 4) = 24.5
myArray(6, 5) = 73.7
myArray(6, 6) = 29
myArray(6, 7) = 244.5
myArray(6, 8) = 344
myArray(6, 9) = 372
myArray(6, 10) = 223
myArray(6, 11) = 34.5
myArray(6, 12) = 42
myArray(6, 13) = 6.5

myArray(7, 1) = 2001
myArray(7, 2) = 63.2
myArray(7, 3) = 76.5
myArray(7, 4) = 17
myArray(7, 5) = 22.5
myArray(7, 6) = 22.5
myArray(7, 7) = 212.5
myArray(7, 8) = 203
myArray(7, 9) = 43
myArray(7, 10) = 87
myArray(7, 11) = 96
myArray(7, 12) = 12
myArray(7, 13) = 24.1

myArray(8, 1) = 2002
myArray(8, 2) = 71.5
myArray(8, 3) = 7.7
myArray(8, 4) = 52
myArray(8, 5) = 149.5
myArray(8, 6) = 127.5
myArray(8, 7) = 57
myArray(8, 8) = 139.5
myArray(8, 9) = 551
myArray(8, 10) = 98.5
myArray(8, 11) = 55.5
myArray(8, 12) = 23.2
myArray(8, 13) = 57.8

myArray(9, 1) = 2003
myArray(9, 2) = 22.4
myArray(9, 3) = 66
myArray(9, 4) = 44
myArray(9, 5) = 202.5
myArray(9, 6) = 164
myArray(9, 7) = 138
myArray(9, 8) = 575
myArray(9, 9) = 280.5
myArray(9, 10) = 192
myArray(9, 11) = 22.5
myArray(9, 12) = 42.5
myArray(9, 13) = 17

myArray(10, 1) = 2004
myArray(10, 2) = 11.2
myArray(10, 3) = 27.3
myArray(10, 4) = 33
myArray(10, 5) = 75.5
myArray(10, 6) = 90.5
myArray(10, 7) = 323.5
myArray(10, 8) = 406
myArray(10, 9) = 330.5
myArray(10, 10) = 126
myArray(10, 11) = 2.5
myArray(10, 12) = 43
myArray(10, 13) = 34.5

myArray(11, 1) = 2005
myArray(11, 2) = 9.4
myArray(11, 3) = 34
myArray(11, 4) = 51
myArray(11, 5) = 31.5
myArray(11, 6) = 65.5
myArray(11, 7) = 191
myArray(11, 8) = 411.5
myArray(11, 9) = 387
myArray(11, 10) = 118
myArray(11, 11) = 23
myArray(11, 12) = 30.5
myArray(11, 13) = 22.6

myArray(12, 1) = 2006
myArray(12, 2) = 28
myArray(12, 3) = 41.1
myArray(12, 4) = 8.4
myArray(12, 5) = 112
myArray(12, 6) = 93.5
myArray(12, 7) = 73
myArray(12, 8) = 681.5
myArray(12, 9) = 118
myArray(12, 10) = 40.5
myArray(12, 11) = 54
myArray(12, 12) = 71
myArray(12, 13) = 28.9

myArray(13, 1) = 2007
myArray(13, 2) = 13.7
myArray(13, 3) = 57
myArray(13, 4) = 129
myArray(13, 5) = 27.5
myArray(13, 6) = 104
myArray(13, 7) = 180
myArray(13, 8) = 252
myArray(13, 9) = 343.5
myArray(13, 10) = 398.5
myArray(13, 11) = 32
myArray(13, 12) = 13.5
myArray(13, 13) = 35.4

myArray(14, 1) = 2008
myArray(14, 2) = 32.4
myArray(14, 3) = 6.1
myArray(14, 4) = 28.3
myArray(14, 5) = 37.6
myArray(14, 6) = 84.5
myArray(14, 7) = 190.5
myArray(14, 8) = 202
myArray(14, 9) = 210
myArray(14, 10) = 35.9
myArray(14, 11) = 40.1
myArray(14, 12) = 15.1
myArray(14, 13) = 19.7

myArray(15, 1) = 2009
myArray(15, 2) = 13.2
myArray(15, 3) = 41.5
myArray(15, 4) = 43
myArray(15, 5) = 36
myArray(15, 6) = 120.3
myArray(15, 7) = 116.3
myArray(15, 8) = 515.5
myArray(15, 9) = 97
myArray(15, 10) = 54.5
myArray(15, 11) = 24
myArray(15, 12) = 29
myArray(15, 13) = 38.3

myArray(16, 1) = 2010
myArray(16, 2) = 33.6
myArray(16, 3) = 74.5
myArray(16, 4) = 83.8
myArray(16, 5) = 73.7
myArray(16, 6) = 114.5
myArray(16, 7) = 62.5
myArray(16, 8) = 278.5
myArray(16, 9) = 495.6
myArray(16, 10) = 110.3
myArray(16, 11) = 20.2
myArray(16, 12) = 20
myArray(16, 13) = 36.5

myArray(17, 1) = 2011
myArray(17, 2) = 2.2
myArray(17, 3) = 63.5
myArray(17, 4) = 21.5
myArray(17, 5) = 132.9
myArray(17, 6) = 130.6
myArray(17, 7) = 237.8
myArray(17, 8) = 571.2
myArray(17, 9) = 403
myArray(17, 10) = 77.8
myArray(17, 11) = 52.2
myArray(17, 12) = 98
myArray(17, 13) = 7.8

myArray(18, 1) = 2012
myArray(18, 2) = 23.7
myArray(18, 3) = 1.1
myArray(18, 4) = 83.6
myArray(18, 5) = 75.9
myArray(18, 6) = 21.7
myArray(18, 7) = 115.7
myArray(18, 8) = 239.2
myArray(18, 9) = 497.5
myArray(18, 10) = 219.5
myArray(18, 11) = 46.6
myArray(18, 12) = 47.3
myArray(18, 13) = 62.7

myArray(19, 1) = 2013
myArray(19, 2) = 37
myArray(19, 3) = 43.8
myArray(19, 4) = 64.6
myArray(19, 5) = 86.4
myArray(19, 6) = 79.5
myArray(19, 7) = 117.7
myArray(19, 8) = 216.9
myArray(19, 9) = 159.5
myArray(19, 10) = 80.8
myArray(19, 11) = 32.6
myArray(19, 12) = 53.9
myArray(19, 13) = 24.1

myArray(20, 1) = 2014
myArray(20, 2) = 4.1
myArray(20, 3) = 2.7
myArray(20, 4) = 97.9
myArray(20, 5) = 88.7
myArray(20, 6) = 26
myArray(20, 7) = 45.6
myArray(20, 8) = 105.8
myArray(20, 9) = 426.4
myArray(20, 10) = 91.2
myArray(20, 11) = 141.2
myArray(20, 12) = 70.1
myArray(20, 13) = 31.3

myArray(21, 1) = 2015
myArray(21, 2) = 37.6
myArray(21, 3) = 23.4
myArray(21, 4) = 46.6
myArray(21, 5) = 93.5
myArray(21, 6) = 29.5
myArray(21, 7) = 143.7
myArray(21, 8) = 162.3
myArray(21, 9) = 83.6
myArray(21, 10) = 18.6
myArray(21, 11) = 93.5
myArray(21, 12) = 109.6
myArray(21, 13) = 35.7

myArray(22, 1) = 2016
myArray(22, 2) = 11.1
myArray(22, 3) = 46
myArray(22, 4) = 54.5
myArray(22, 5) = 171.7
myArray(22, 6) = 70.5
myArray(22, 7) = 87.4
myArray(22, 8) = 377.9
myArray(22, 9) = 105.6
myArray(22, 10) = 160.9
myArray(22, 11) = 157.2
myArray(22, 12) = 33.2
myArray(22, 13) = 49.6

myArray(23, 1) = 2017
myArray(23, 2) = 13.6
myArray(23, 3) = 54.6
myArray(23, 4) = 29.8
myArray(23, 5) = 76.1
myArray(23, 6) = 31.8
myArray(23, 7) = 48.3
myArray(23, 8) = 305.5
myArray(23, 9) = 222.3
myArray(23, 10) = 105.6
myArray(23, 11) = 35.1
myArray(23, 12) = 15.6
myArray(23, 13) = 29.3

myArray(24, 1) = 2018
myArray(24, 2) = 25.7
myArray(24, 3) = 28.1
myArray(24, 4) = 91.5
myArray(24, 5) = 142.4
myArray(24, 6) = 110.4
myArray(24, 7) = 104.3
myArray(24, 8) = 163.5
myArray(24, 9) = 410.4
myArray(24, 10) = 135.2
myArray(24, 11) = 112.6
myArray(24, 12) = 45.5
myArray(24, 13) = 27.6

myArray(25, 1) = 2019
myArray(25, 2) = 6.4
myArray(25, 3) = 41.5
myArray(25, 4) = 33
myArray(25, 5) = 93
myArray(25, 6) = 44.2
myArray(25, 7) = 101
myArray(25, 8) = 141.1
myArray(25, 9) = 105.8
myArray(25, 10) = 236.4
myArray(25, 11) = 99.3
myArray(25, 12) = 47.9
myArray(25, 13) = 33

myArray(26, 1) = 2020
myArray(26, 2) = 80.8
myArray(26, 3) = 83.9
myArray(26, 4) = 20.5
myArray(26, 5) = 35.6
myArray(26, 6) = 80.5
myArray(26, 7) = 234
myArray(26, 8) = 628
myArray(26, 9) = 373.4
myArray(26, 10) = 167.2
myArray(26, 11) = 4.1
myArray(26, 12) = 41.9
myArray(26, 13) = 8.3

myArray(27, 1) = 2021
myArray(27, 2) = 23.5
myArray(27, 3) = 19.3
myArray(27, 4) = 88
myArray(27, 5) = 39.3
myArray(27, 6) = 162.7
myArray(27, 7) = 105.6
myArray(27, 8) = 300.8
myArray(27, 9) = 297.2
myArray(27, 10) = 151.9
myArray(27, 11) = 44
myArray(27, 12) = 50.7
myArray(27, 13) = 7.1

myArray(28, 1) = 2022
myArray(28, 2) = 1.5
myArray(28, 3) = 4.1
myArray(28, 4) = 80.6
myArray(28, 5) = 63.3
myArray(28, 6) = 4.7
myArray(28, 7) = 145.4
myArray(28, 8) = 183.7
myArray(28, 9) = 265.7
myArray(28, 10) = 68.2
myArray(28, 11) = 59.3
myArray(28, 12) = 54.2
myArray(28, 13) = 18

myArray(29, 1) = 2023
myArray(29, 2) = 28.7
myArray(29, 3) = 8.7
myArray(29, 4) = 22
myArray(29, 5) = 46.1
myArray(29, 6) = 211.5
myArray(29, 7) = 196.9
myArray(29, 8) = 624.5
myArray(29, 9) = 248.7
myArray(29, 10) = 218.5
myArray(29, 11) = 7.6
myArray(29, 12) = 60
myArray(29, 13) = 127.7

myArray(30, 1) = 2024
myArray(30, 2) = 54.3
myArray(30, 3) = 95.7
myArray(30, 4) = 72.7
myArray(30, 5) = 46.4
myArray(30, 6) = 83.2
myArray(30, 7) = 119.7
myArray(30, 8) = 501.5
myArray(30, 9) = 56.7
myArray(30, 10) = 264.8
myArray(30, 11) = 81
myArray(30, 12) = 32.7
myArray(30, 13) = 7.5


    data_GEUMSAN = myArray

End Function











Sub ChangeFileNameInCurrentDir()
    Dim CurrentDir As String
    Dim OldFileName As String
    Dim NewFileName As String

    ' Get the current working directory
    CurrentDir = ThisWorkbook.Path & "\"

    ' Define the old and new file names
    OldFileName = "myArray.csv"
    NewFileName = ActiveSheet.name & ".csv"

    ' Check if the old file exists in the current directory
    If Dir(CurrentDir & OldFileName) <> "" Then
        ' Rename the file
        Name CurrentDir & OldFileName As CurrentDir & NewFileName
        MsgBox "File name changed successfully!"
    Else
        MsgBox "The old file does not exist in the current directory."
    End If
End Sub


Sub SaveRangeToFile()
    Dim ws As Worksheet
    Dim rng As Range
    Dim filePath As String
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("B6:N35")
    
    ' Generate the file path using the worksheet name
    filePath = ThisWorkbook.Path & "\" & ws.name & ".csv"
    
    ' Save the range to a CSV file
    rng.ExportAsFixedFormat Type:=xlCSV, fileName:=filePath, Quality:=xlQualityStandard
End Sub











'2024-01-02
'이것이 파일로 세이브 하는 메인함수이다.
'이것으로 강수량 데이타를 세이브 할수있다.

Sub DumpRangeToArrayAndSaveTest()
' Ctrl+D 로 세이브 해주는 함수

    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    Dim AREA_STR As String
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.Value
    
    ' Save array to a file
    Dim filePath As String
    
    
    filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
    SaveArrayToFileByExcelForm myArray, filePath
 
End Sub


Function getAreaName()
    Dim lookupValue As Variant
    Dim result As Variant

    Dim tableRange As Range
    Set tableRange = Range("tblAREAREF")

    If ActiveSheet.name = "main" Then
        lookupValue = Range("S8")
    Else
        lookupValue = ActiveSheet.name
    End If
    
    On Error Resume Next
    result = Application.VLookup(lookupValue, tableRange, 2, False)
    On Error GoTo 0
    
    If Not IsError(result) Then
        getAreaName = UCase(result)
    Else
        ' If no match is found, display an error message
        getAreaName = "MAIN"
    End If

End Function



Private Sub SaveArrayToFileByExcelForm(myArray As Variant, filePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    Dim AREA_STR As String
    
    FileNum = FreeFile
    AREA_STR = getAreaName()
    
    
    Open filePath For Output As FileNum
    
    Print #FileNum, "Function data_" & AREA_STR & "() As Variant"
    Print #FileNum, ""
    Print #FileNum, "    Dim myArray() As Variant"
    Print #FileNum, "    ReDim myArray(1 To 30, 1 To 13)"
    Print #FileNum, ""
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Print #FileNum, "myArray(" & i & ", " & j & ") = ";
            
            ' Separate values with a comma (CSV format)
            If j <= UBound(myArray, 2) Then
                Print #FileNum, myArray(i, j);
            End If
            
            Print #FileNum, ""
        Next j
        ' Start a new line for each row
        Print #FileNum, ""
    Next i
    
    Print #FileNum, ""
    Print #FileNum, "    data_" & AREA_STR & "= myArray"
    Print #FileNum, ""
    Print #FileNum, "End Function"
    
    Close FileNum
End Sub


Sub importFromArray()
    Dim myArray As Variant
    Dim rng As Range
    
    indexString = "data_" & UCase(Range("s11").Value)
    
    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0
    
    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If
    
    
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    rng.Value = myArray
       
End Sub





Option Explicit

'
' Sheet1(AREAREF)
'

Option Explicit



Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    Dim fName As String
    
    Dim pathToExport As String
    pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        
        
        If item.CodeModule.CountOfLines <> 0 Then
            lineToPrint = item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        Else
            lineToPrint = "'This Module is Empty "
        End If
        
        
        fName = item.CodeModule.name
        Debug.Print lineToPrint
        SaveTextToFile lineToPrint, pathToExport & fName & ".bas"
        
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub


Function Is64BitOS() As Boolean
    Dim osInfo As String
    osInfo = Application.OperatingSystem
    
    If InStr(osInfo, "64") > 0 Then
        Is64BitOS = True
    Else
        Is64BitOS = False
    End If
End Function


Sub OpenSpecificChromeVersion()
  Const chromePath = "c:\ProgramData\00_chrome\chrome.exe"
  Dim objShell As Object
  Set objShell = CreateObject("Wscript.Shell")

  objShell.Run Chr(34) & chromePath & Chr(34)
  Set objShell = Nothing
End Sub


Sub RunChrome_SpecificLocationOfChrome()
    Dim driver As New Selenium.chromeDriver
    Dim url As String
    
    ' 2024/12/26
    ' this is the key of run specific version of chrome
    ' driver.SetBinary "c:\ProgramData\00_chrome\chrome.exe" ' Update this path

    If Is64BitOS() Then
        Set SystemProperties = CreateObject("WScript.Shell").Environment("System")
        SystemProperties("webdriver.chrome.driver") = driverPath  ' For 64-bit systems
    Else
        SystemProperties("webdriver.chrome.driver") = driverPath  ' For 32-bit systems
    End If
    
    driver.AddArgument "--remote-debugging-port=9222"  ' Specify the port number
    'driver.AddArgument "--user-data-dir=C:\Users\minhwasoo\AppData\Local\Google\Chrome\User Data"  ' Specify the user data directory
    driver.AddArgument "--disable-gpu"  ' Disable GPU acceleration
    driver.AddArgument "--window-size=1920,1080"  ' Specify the window size
        
    driver.SetBinary "c:\ProgramData\00_chrome\chrome.exe" ' Update this path
    url = "https://www.google.com" ' Update this URL
   
    driver.Start
    driver.Get url
  
    Sleep (7 * 1000)
    driver.Quit
End Sub


Sub RunChromeWithDownloadDirectory()
   
   ' https://stackoverflow.com/questions/63424206/selenium-chromedriver-excel-vba-change-the-download-directory-multiple-time
    
    Dim driver As New Selenium.chromeDriver
    
    driver.SetPreference "download.default_directory", "C:\Folder\"
    driver.SetPreference "download.directory_upgrade", True
    driver.SetPreference "download.prompt_for_download", False

    driver.Get "https://www.google.com/"

End Sub




   





Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Private Sub CommandButton_AnnualReset_Click()
    Dim ws As Worksheet
    Dim sheetNames As String
        
    DeleteSheets

End Sub

Sub DeleteSheets()
    Dim ws As Worksheet
    Dim sheetNamesToKeep As Variant
    
    ' Define the sheet names to keep
    sheetNamesToKeep = Array("main", "AREA", "AREAREF")
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet is not in the list of sheet names to keep
        If IsError(Application.Match(ws.name, sheetNamesToKeep, 0)) Then
            ' Delete the sheet
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
End Sub


Private Sub CommandButton_BackUP_Click()
    Call BackupData
    Sheets("main").Activate
End Sub

Private Sub CommandButton_Clear30Year_Click()
    Call clear_30year_data
End Sub


Private Sub CommandButton_DeleteIgnoreError_Click()
    Call deleteall_igonre_error
End Sub

Private Sub CommandButton_GetWeatherData_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   Call deleteall_igonre_error
   
End Sub

Private Sub CommandButton_LoadDataFromArray_Click()
   Call modDumpArrayMyData.importFromArray
End Sub
Function data_SEOUL() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

myArray(1, 1) = 1995
myArray(1, 2) = 11.6
myArray(1, 3) = 5.2
myArray(1, 4) = 60.6
myArray(1, 5) = 44.4
myArray(1, 6) = 60.6
myArray(1, 7) = 70.7
myArray(1, 8) = 436.1
myArray(1, 9) = 786.6
myArray(1, 10) = 47.2
myArray(1, 11) = 39.3
myArray(1, 12) = 32.9
myArray(1, 13) = 3.4

myArray(2, 1) = 1996
myArray(2, 2) = 16.3
myArray(2, 3) = 1
myArray(2, 4) = 77.9
myArray(2, 5) = 62
myArray(2, 6) = 29.3
myArray(2, 7) = 249.7
myArray(2, 8) = 512.8
myArray(2, 9) = 132.4
myArray(2, 10) = 11
myArray(2, 11) = 90.3
myArray(2, 12) = 62.9
myArray(2, 13) = 11

myArray(3, 1) = 1997
myArray(3, 2) = 16.8
myArray(3, 3) = 39.6
myArray(3, 4) = 25.3
myArray(3, 5) = 56.1
myArray(3, 6) = 291.3
myArray(3, 7) = 110
myArray(3, 8) = 299.6
myArray(3, 9) = 117.2
myArray(3, 10) = 76.9
myArray(3, 11) = 45.5
myArray(3, 12) = 93.8
myArray(3, 13) = 38.1

myArray(4, 1) = 1998
myArray(4, 2) = 10.4
myArray(4, 3) = 32.3
myArray(4, 4) = 45.1
myArray(4, 5) = 120.2
myArray(4, 6) = 121.5
myArray(4, 7) = 234.1
myArray(4, 8) = 311.8
myArray(4, 9) = 1237.8
myArray(4, 10) = 177.9
myArray(4, 11) = 27.4
myArray(4, 12) = 26.9
myArray(4, 13) = 3.7

myArray(5, 1) = 1999
myArray(5, 2) = 10.2
myArray(5, 3) = 2.9
myArray(5, 4) = 55
myArray(5, 5) = 97.2
myArray(5, 6) = 109.7
myArray(5, 7) = 131.8
myArray(5, 8) = 230.4
myArray(5, 9) = 600.5
myArray(5, 10) = 377.3
myArray(5, 11) = 81.6
myArray(5, 12) = 19.5
myArray(5, 13) = 17

myArray(6, 1) = 2000
myArray(6, 2) = 42.8
myArray(6, 3) = 2.1
myArray(6, 4) = 3.1
myArray(6, 5) = 30.7
myArray(6, 6) = 75.2
myArray(6, 7) = 68.1
myArray(6, 8) = 114.7
myArray(6, 9) = 599.4
myArray(6, 10) = 178.5
myArray(6, 11) = 18.1
myArray(6, 12) = 27.1
myArray(6, 13) = 27

myArray(7, 1) = 2001
myArray(7, 2) = 39.4
myArray(7, 3) = 45.7
myArray(7, 4) = 18.1
myArray(7, 5) = 12.3
myArray(7, 6) = 16.5
myArray(7, 7) = 157.4
myArray(7, 8) = 698.4
myArray(7, 9) = 252
myArray(7, 10) = 49.3
myArray(7, 11) = 68.2
myArray(7, 12) = 13
myArray(7, 13) = 15.7

myArray(8, 1) = 2002
myArray(8, 2) = 37.4
myArray(8, 3) = 2.4
myArray(8, 4) = 31.5
myArray(8, 5) = 155.1
myArray(8, 6) = 58
myArray(8, 7) = 61.4
myArray(8, 8) = 220.6
myArray(8, 9) = 688
myArray(8, 10) = 61.1
myArray(8, 11) = 45
myArray(8, 12) = 12.5
myArray(8, 13) = 15

myArray(9, 1) = 2003
myArray(9, 2) = 14.1
myArray(9, 3) = 39.6
myArray(9, 4) = 26.8
myArray(9, 5) = 139.6
myArray(9, 6) = 106
myArray(9, 7) = 156
myArray(9, 8) = 469.8
myArray(9, 9) = 684.2
myArray(9, 10) = 258.2
myArray(9, 11) = 41.5
myArray(9, 12) = 69.3
myArray(9, 13) = 6.9

myArray(10, 1) = 2004
myArray(10, 2) = 19.8
myArray(10, 3) = 54.6
myArray(10, 4) = 27.6
myArray(10, 5) = 74.1
myArray(10, 6) = 168.5
myArray(10, 7) = 138.1
myArray(10, 8) = 510.7
myArray(10, 9) = 193.3
myArray(10, 10) = 198.7
myArray(10, 11) = 6.5
myArray(10, 12) = 80
myArray(10, 13) = 27.2

myArray(11, 1) = 2005
myArray(11, 2) = 4.5
myArray(11, 3) = 17.2
myArray(11, 4) = 12.5
myArray(11, 5) = 94.7
myArray(11, 6) = 85.8
myArray(11, 7) = 168.5
myArray(11, 8) = 269.4
myArray(11, 9) = 285
myArray(11, 10) = 313.3
myArray(11, 11) = 52.6
myArray(11, 12) = 44.6
myArray(11, 13) = 10.3

myArray(12, 1) = 2006
myArray(12, 2) = 34.3
myArray(12, 3) = 15.7
myArray(12, 4) = 14
myArray(12, 5) = 51.8
myArray(12, 6) = 156.2
myArray(12, 7) = 168.5
myArray(12, 8) = 1014
myArray(12, 9) = 121.2
myArray(12, 10) = 11.1
myArray(12, 11) = 30.2
myArray(12, 12) = 47.6
myArray(12, 13) = 17.3

myArray(13, 1) = 2007
myArray(13, 2) = 10.8
myArray(13, 3) = 12.6
myArray(13, 4) = 123.5
myArray(13, 5) = 41.1
myArray(13, 6) = 137.6
myArray(13, 7) = 54.5
myArray(13, 8) = 274.1
myArray(13, 9) = 237.6
myArray(13, 10) = 241.9
myArray(13, 11) = 39.5
myArray(13, 12) = 26.4
myArray(13, 13) = 12.7

myArray(14, 1) = 2008
myArray(14, 2) = 17.7
myArray(14, 3) = 15
myArray(14, 4) = 53.9
myArray(14, 5) = 38.5
myArray(14, 6) = 97.7
myArray(14, 7) = 165
myArray(14, 8) = 530.8
myArray(14, 9) = 251.2
myArray(14, 10) = 99.2
myArray(14, 11) = 41.8
myArray(14, 12) = 19.6
myArray(14, 13) = 25.9

myArray(15, 1) = 2009
myArray(15, 2) = 5.7
myArray(15, 3) = 36.9
myArray(15, 4) = 63.9
myArray(15, 5) = 66.5
myArray(15, 6) = 109
myArray(15, 7) = 132
myArray(15, 8) = 659.4
myArray(15, 9) = 285.3
myArray(15, 10) = 64.5
myArray(15, 11) = 66.9
myArray(15, 12) = 52.4
myArray(15, 13) = 21.5

myArray(16, 1) = 2010
myArray(16, 2) = 29.3
myArray(16, 3) = 55.3
myArray(16, 4) = 82.5
myArray(16, 5) = 62.8
myArray(16, 6) = 124
myArray(16, 7) = 127.6
myArray(16, 8) = 239.2
myArray(16, 9) = 598.7
myArray(16, 10) = 671.5
myArray(16, 11) = 25.6
myArray(16, 12) = 10.9
myArray(16, 13) = 16.1

myArray(17, 1) = 2011
myArray(17, 2) = 8.9
myArray(17, 3) = 29.1
myArray(17, 4) = 14.6
myArray(17, 5) = 110.1
myArray(17, 6) = 53.4
myArray(17, 7) = 404.5
myArray(17, 8) = 1131
myArray(17, 9) = 166.8
myArray(17, 10) = 25.6
myArray(17, 11) = 32
myArray(17, 12) = 56.2
myArray(17, 13) = 7.1

myArray(18, 1) = 2012
myArray(18, 2) = 6.7
myArray(18, 3) = 0.8
myArray(18, 4) = 47.4
myArray(18, 5) = 157
myArray(18, 6) = 8.2
myArray(18, 7) = 91.9
myArray(18, 8) = 448.9
myArray(18, 9) = 464.9
myArray(18, 10) = 212
myArray(18, 11) = 99.3
myArray(18, 12) = 67.8
myArray(18, 13) = 41.4

myArray(19, 1) = 2013
myArray(19, 2) = 22.1
myArray(19, 3) = 74.1
myArray(19, 4) = 27.3
myArray(19, 5) = 71.7
myArray(19, 6) = 132
myArray(19, 7) = 28.3
myArray(19, 8) = 676.2
myArray(19, 9) = 148.6
myArray(19, 10) = 138.5
myArray(19, 11) = 13.5
myArray(19, 12) = 46.8
myArray(19, 13) = 24.7

myArray(20, 1) = 2014
myArray(20, 2) = 13
myArray(20, 3) = 16.2
myArray(20, 4) = 7.2
myArray(20, 5) = 31
myArray(20, 6) = 63
myArray(20, 7) = 98.1
myArray(20, 8) = 207.9
myArray(20, 9) = 172.8
myArray(20, 10) = 88.1
myArray(20, 11) = 52.2
myArray(20, 12) = 41.5
myArray(20, 13) = 17.9

myArray(21, 1) = 2015
myArray(21, 2) = 11.3
myArray(21, 3) = 22.7
myArray(21, 4) = 9.6
myArray(21, 5) = 80.5
myArray(21, 6) = 28.9
myArray(21, 7) = 99
myArray(21, 8) = 226
myArray(21, 9) = 72.9
myArray(21, 10) = 26
myArray(21, 11) = 81.5
myArray(21, 12) = 104.6
myArray(21, 13) = 29.1

myArray(22, 1) = 2016
myArray(22, 2) = 1
myArray(22, 3) = 47.6
myArray(22, 4) = 40.5
myArray(22, 5) = 76.8
myArray(22, 6) = 160.5
myArray(22, 7) = 54.4
myArray(22, 8) = 358.2
myArray(22, 9) = 67.1
myArray(22, 10) = 33
myArray(22, 11) = 74.8
myArray(22, 12) = 16.7
myArray(22, 13) = 61.1

myArray(23, 1) = 2017
myArray(23, 2) = 14.9
myArray(23, 3) = 11.1
myArray(23, 4) = 7.9
myArray(23, 5) = 61.6
myArray(23, 6) = 16.1
myArray(23, 7) = 66.6
myArray(23, 8) = 621
myArray(23, 9) = 297
myArray(23, 10) = 35
myArray(23, 11) = 26.5
myArray(23, 12) = 40.7
myArray(23, 13) = 34.8

myArray(24, 1) = 2018
myArray(24, 2) = 8.5
myArray(24, 3) = 29.6
myArray(24, 4) = 49.5
myArray(24, 5) = 130.3
myArray(24, 6) = 222
myArray(24, 7) = 171.5
myArray(24, 8) = 185.6
myArray(24, 9) = 202.6
myArray(24, 10) = 68.5
myArray(24, 11) = 120.5
myArray(24, 12) = 79.1
myArray(24, 13) = 16.4

myArray(25, 1) = 2019
myArray(25, 2) = 0
myArray(25, 3) = 23.8
myArray(25, 4) = 26.8
myArray(25, 5) = 47.3
myArray(25, 6) = 37.8
myArray(25, 7) = 74
myArray(25, 8) = 194.4
myArray(25, 9) = 190.5
myArray(25, 10) = 139.8
myArray(25, 11) = 55.5
myArray(25, 12) = 78.8
myArray(25, 13) = 22.6

myArray(26, 1) = 2020
myArray(26, 2) = 60.5
myArray(26, 3) = 53.1
myArray(26, 4) = 16.3
myArray(26, 5) = 16.9
myArray(26, 6) = 112.4
myArray(26, 7) = 139.6
myArray(26, 8) = 270.4
myArray(26, 9) = 675.7
myArray(26, 10) = 181.5
myArray(26, 11) = 0
myArray(26, 12) = 120.1
myArray(26, 13) = 4.6

myArray(27, 1) = 2021
myArray(27, 2) = 18.9
myArray(27, 3) = 7.1
myArray(27, 4) = 110.9
myArray(27, 5) = 124.1
myArray(27, 6) = 183.1
myArray(27, 7) = 104.6
myArray(27, 8) = 168.3
myArray(27, 9) = 211.2
myArray(27, 10) = 131
myArray(27, 11) = 57
myArray(27, 12) = 62.4
myArray(27, 13) = 7.9

myArray(28, 1) = 2022
myArray(28, 2) = 5.5
myArray(28, 3) = 4.7
myArray(28, 4) = 102.6
myArray(28, 5) = 20.4
myArray(28, 6) = 7.5
myArray(28, 7) = 393.8
myArray(28, 8) = 252.3
myArray(28, 9) = 564.8
myArray(28, 10) = 201.5
myArray(28, 11) = 124.1
myArray(28, 12) = 84.5
myArray(28, 13) = 13.6

myArray(29, 1) = 2023
myArray(29, 2) = 47.9
myArray(29, 3) = 1
myArray(29, 4) = 10.5
myArray(29, 5) = 96.9
myArray(29, 6) = 155.6
myArray(29, 7) = 195.6
myArray(29, 8) = 459.9
myArray(29, 9) = 298.1
myArray(29, 10) = 134.5
myArray(29, 11) = 31
myArray(29, 12) = 81.9
myArray(29, 13) = 85.9

myArray(30, 1) = 2024
myArray(30, 2) = 18.9
myArray(30, 3) = 74.7
myArray(30, 4) = 29.9
myArray(30, 5) = 33.2
myArray(30, 6) = 125.1
myArray(30, 7) = 115.9
myArray(30, 8) = 557.3
myArray(30, 9) = 72.8
myArray(30, 10) = 143.9
myArray(30, 11) = 74
myArray(30, 12) = 60
myArray(30, 13) = 5.7


    data_SEOUL = myArray

End Function

