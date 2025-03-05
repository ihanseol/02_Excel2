Function GetProvince_Dictionary(city As String) As String
    Dim cityProvinceMap As Object
    Set cityProvinceMap = CreateObject("Scripting.Dictionary")
    
    ' 충청도
    cityProvinceMap.Add "보은", "충청도"
    cityProvinceMap.Add "제천", "충청도"
    cityProvinceMap.Add "청주", "충청도"
    cityProvinceMap.Add "추풍령", "충청도"
    cityProvinceMap.Add "대전", "충청도"
    cityProvinceMap.Add "세종", "충청도"
    cityProvinceMap.Add "금산", "충청도"
    cityProvinceMap.Add "보령", "충청도"
    cityProvinceMap.Add "부여", "충청도"
    cityProvinceMap.Add "서산", "충청도"
    cityProvinceMap.Add "천안", "충청도"
    cityProvinceMap.Add "홍성", "충청도"

    ' 서울경기
    cityProvinceMap.Add "관악산", "서울경기"
    cityProvinceMap.Add "서울", "서울경기"
    cityProvinceMap.Add "강화", "서울경기"
    cityProvinceMap.Add "백령도", "서울경기"
    cityProvinceMap.Add "인천", "서울경기"
    cityProvinceMap.Add "동두천", "서울경기"
    cityProvinceMap.Add "수원", "서울경기"
    cityProvinceMap.Add "양평", "서울경기"
    cityProvinceMap.Add "이천", "서울경기"
    cityProvinceMap.Add "파주", "서울경기"

    ' 강원도
    cityProvinceMap.Add "강릉", "강원도"
    cityProvinceMap.Add "대관령", "강원도"
    cityProvinceMap.Add "동해", "강원도"
    cityProvinceMap.Add "북강릉", "강원도"
    cityProvinceMap.Add "북춘천", "강원도"
    cityProvinceMap.Add "삼척", "강원도"
    cityProvinceMap.Add "속초", "강원도"
    cityProvinceMap.Add "영월", "강원도"
    cityProvinceMap.Add "원주", "강원도"
    cityProvinceMap.Add "인제", "강원도"
    cityProvinceMap.Add "정선군", "강원도"
    cityProvinceMap.Add "철원", "강원도"
    cityProvinceMap.Add "춘천", "강원도"
    cityProvinceMap.Add "태백", "강원도"
    cityProvinceMap.Add "홍천", "강원도"

    ' 전라도
    cityProvinceMap.Add "광주", "전라도"
    cityProvinceMap.Add "고창", "전라도"
    cityProvinceMap.Add "고창군", "전라도"
    cityProvinceMap.Add "군산", "전라도"
    cityProvinceMap.Add "남원", "전라도"
    cityProvinceMap.Add "부안", "전라도"
    cityProvinceMap.Add "순창군", "전라도"
    cityProvinceMap.Add "임실", "전라도"
    cityProvinceMap.Add "장수", "전라도"
    cityProvinceMap.Add "전주", "전라도"
    cityProvinceMap.Add "정읍", "전라도"
    cityProvinceMap.Add "강진군", "전라도"
    cityProvinceMap.Add "고흥", "전라도"
    cityProvinceMap.Add "광양시", "전라도"
    cityProvinceMap.Add "목포", "전라도"
    cityProvinceMap.Add "무안", "전라도"
    cityProvinceMap.Add "보성군", "전라도"
    cityProvinceMap.Add "순천", "전라도"
    cityProvinceMap.Add "여수", "전라도"
    cityProvinceMap.Add "영광군", "전라도"
    cityProvinceMap.Add "완도", "전라도"
    cityProvinceMap.Add "장흥", "전라도"
    cityProvinceMap.Add "주암", "전라도"
    cityProvinceMap.Add "진도", "전라도"
    cityProvinceMap.Add "첨철산", "전라도"
    cityProvinceMap.Add "진도군", "전라도"
    cityProvinceMap.Add "해남", "전라도"
    cityProvinceMap.Add "흑산도", "전라도"

    ' 경상도
    cityProvinceMap.Add "대구", "경상도"
    cityProvinceMap.Add "대구(기)", "경상도"
    cityProvinceMap.Add "울산", "경상도"
    cityProvinceMap.Add "부산", "경상도"
    cityProvinceMap.Add "경주시", "경상도"
    cityProvinceMap.Add "구미", "경상도"
    cityProvinceMap.Add "문경", "경상도"
    cityProvinceMap.Add "봉화", "경상도"
    cityProvinceMap.Add "상주", "경상도"
    cityProvinceMap.Add "안동", "경상도"
    cityProvinceMap.Add "영덕", "경상도"
    cityProvinceMap.Add "영주", "경상도"
    cityProvinceMap.Add "영천", "경상도"
    cityProvinceMap.Add "울릉도", "경상도"
    cityProvinceMap.Add "울진", "경상도"
    cityProvinceMap.Add "의성", "경상도"
    cityProvinceMap.Add "청송군", "경상도"
    cityProvinceMap.Add "포항", "경상도"
    cityProvinceMap.Add "거제", "경상도"
    cityProvinceMap.Add "거창", "경상도"
    cityProvinceMap.Add "김해시", "경상도"
    cityProvinceMap.Add "남해", "경상도"
    cityProvinceMap.Add "밀양", "경상도"
    cityProvinceMap.Add "북창원", "경상도"
    cityProvinceMap.Add "산청", "경상도"
    cityProvinceMap.Add "양산시", "경상도"
    cityProvinceMap.Add "의령군", "경상도"
    cityProvinceMap.Add "진주", "경상도"
    cityProvinceMap.Add "창원", "경상도"
    cityProvinceMap.Add "통영", "경상도"
    cityProvinceMap.Add "함양군", "경상도"
    cityProvinceMap.Add "합천", "경상도"

    ' 제주도
    cityProvinceMap.Add "고산", "제주도"
    cityProvinceMap.Add "서귀포", "제주도"
    cityProvinceMap.Add "성산", "제주도"
    cityProvinceMap.Add "성산2", "제주도"
    cityProvinceMap.Add "성산포", "제주도"
    
    ' Return the province if found
    If cityProvinceMap.Exists(city) Then
        GetProvince_Dictionary = cityProvinceMap(city)
    Else
        GetProvince_Dictionary = "Not in list"
    End If
    
    ' Clean up
    Set cityProvinceMap = Nothing
End Function



Function GetProvince_Vlookup(city As String) As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim result As Variant
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to the name of your sheet
    Set rng = ws.Range("A1:B100") ' Adjust the range to match your data

    ' Use VLOOKUP to find the province
    result = Application.WorksheetFunction.VLookup(city, rng, 2, False)
    
    ' Return the result or handle the error
    If IsError(result) Then
        GetProvince_Vlookup = "Not in list"
    Else
        GetProvince_Vlookup = result
    End If
End Function


Function GetProvince_Collection(city As String) As String
    Dim cityProvinceArray As Collection
    Dim cityProvince As Variant
    Dim i As Integer
    
    ' Initialize the collection with city-province pairs
    Set cityProvinceArray = New Collection
    
    ' Add city-province pairs to the collection
    cityProvinceArray.Add Array("보은", "충청도")
    cityProvinceArray.Add Array("제천", "충청도")
    cityProvinceArray.Add Array("청주", "충청도")
    cityProvinceArray.Add Array("추풍령", "충청도")
    cityProvinceArray.Add Array("대전", "충청도")
    cityProvinceArray.Add Array("세종", "충청도")
    cityProvinceArray.Add Array("금산", "충청도")
    cityProvinceArray.Add Array("보령", "충청도")
    cityProvinceArray.Add Array("부여", "충청도")
    cityProvinceArray.Add Array("서산", "충청도")
    cityProvinceArray.Add Array("천안", "충청도")
    cityProvinceArray.Add Array("홍성", "충청도")
    cityProvinceArray.Add Array("관악산", "서울경기")
    cityProvinceArray.Add Array("서울", "서울경기")
    cityProvinceArray.Add Array("강화", "서울경기")
    cityProvinceArray.Add Array("백령도", "서울경기")
    cityProvinceArray.Add Array("인천", "서울경기")
    cityProvinceArray.Add Array("동두천", "서울경기")
    cityProvinceArray.Add Array("수원", "서울경기")
    cityProvinceArray.Add Array("양평", "서울경기")
    cityProvinceArray.Add Array("이천", "서울경기")
    cityProvinceArray.Add Array("파주", "서울경기")
    cityProvinceArray.Add Array("강릉", "강원도")
    cityProvinceArray.Add Array("대관령", "강원도")
    cityProvinceArray.Add Array("동해", "강원도")
    cityProvinceArray.Add Array("북강릉", "강원도")
    cityProvinceArray.Add Array("북춘천", "강원도")
    cityProvinceArray.Add Array("삼척", "강원도")
    cityProvinceArray.Add Array("속초", "강원도")
    cityProvinceArray.Add Array("영월", "강원도")
    cityProvinceArray.Add Array("원주", "강원도")
    cityProvinceArray.Add Array("인제", "강원도")
    cityProvinceArray.Add Array("정선군", "강원도")
    cityProvinceArray.Add Array("철원", "강원도")
    cityProvinceArray.Add Array("춘천", "강원도")
    cityProvinceArray.Add Array("태백", "강원도")
    cityProvinceArray.Add Array("홍천", "강원도")
    cityProvinceArray.Add Array("광주", "전라도")
    cityProvinceArray.Add Array("고창", "전라도")
    cityProvinceArray.Add Array("고창군", "전라도")
    cityProvinceArray.Add Array("군산", "전라도")
    cityProvinceArray.Add Array("남원", "전라도")
    cityProvinceArray.Add Array("부안", "전라도")
    cityProvinceArray.Add Array("순창군", "전라도")
    cityProvinceArray.Add Array("임실", "전라도")
    cityProvinceArray.Add Array("장수", "전라도")
    cityProvinceArray.Add Array("전주", "전라도")
    cityProvinceArray.Add Array("정읍", "전라도")
    cityProvinceArray.Add Array("강진군", "전라도")
    cityProvinceArray.Add Array("고흥", "전라도")
    cityProvinceArray.Add Array("광양시", "전라도")
    cityProvinceArray.Add Array("목포", "전라도")
    cityProvinceArray.Add Array("무안", "전라도")
    cityProvinceArray.Add Array("보성군", "전라도")
    cityProvinceArray.Add Array("순천", "전라도")
    cityProvinceArray.Add Array("여수", "전라도")
    cityProvinceArray.Add Array("영광군", "전라도")
    cityProvinceArray.Add Array("완도", "전라도")
    cityProvinceArray.Add Array("장흥", "전라도")
    cityProvinceArray.Add Array("주암", "전라도")
    cityProvinceArray.Add Array("진도", "전라도")
    cityProvinceArray.Add Array("첨철산", "전라도")
    cityProvinceArray.Add Array("진도군", "전라도")
    cityProvinceArray.Add Array("해남", "전라도")
    cityProvinceArray.Add Array("흑산도", "전라도")
    cityProvinceArray.Add Array("대구", "경상도")
    cityProvinceArray.Add Array("대구(기)", "경상도")
    cityProvinceArray.Add Array("울산", "경상도")
    cityProvinceArray.Add Array("부산", "경상도")
    cityProvinceArray.Add Array("경주시", "경상도")
    cityProvinceArray.Add Array("구미", "경상도")
    cityProvinceArray.Add Array("문경", "경상도")
    cityProvinceArray.Add Array("봉화", "경상도")
    cityProvinceArray.Add Array("상주", "경상도")
    cityProvinceArray.Add Array("안동", "경상도")
    cityProvinceArray.Add Array("영덕", "경상도")
    cityProvinceArray.Add Array("영주", "경상도")
    cityProvinceArray.Add Array("영천", "경상도")
    cityProvinceArray.Add Array("울릉도", "경상도")
    cityProvinceArray.Add Array("울진", "경상도")
    cityProvinceArray.Add Array("의성", "경상도")
    cityProvinceArray.Add Array("청송군", "경상도")
    cityProvinceArray.Add Array("포항", "경상도")
    cityProvinceArray.Add Array("거제", "경상도")
    cityProvinceArray.Add Array("거창", "경상도")
    cityProvinceArray.Add Array("김해시", "경상도")
    cityProvinceArray.Add Array("남해", "경상도")
    cityProvinceArray.Add Array("밀양", "경상도")
    cityProvinceArray.Add Array("북창원", "경상도")
    cityProvinceArray.Add Array("산청", "경상도")
    cityProvinceArray.Add Array("양산시", "경상도")
    cityProvinceArray.Add Array("의령군", "경상도")
    cityProvinceArray.Add Array("진주", "경상도")
    cityProvinceArray.Add Array("창원", "경상도")
    cityProvinceArray.Add Array("통영", "경상도")
    cityProvinceArray.Add Array("함양군", "경상도")
    cityProvinceArray.Add Array("합천", "경상도")
    cityProvinceArray.Add Array("고산", "제주도")
    cityProvinceArray.Add Array("서귀포", "제주도")
    cityProvinceArray.Add Array("성산", "제주도")
    cityProvinceArray.Add Array("성산2", "제주도")
    cityProvinceArray.Add Array("성산포", "제주도")
    
    ' Loop through the collection to find the city and return the corresponding province
    For Each cityProvince In cityProvinceArray
        If cityProvince(0) = city Then
            GetProvince_Collection = cityProvince(1)
            Exit Function
        End If
    Next cityProvince
    
    ' If city is not found in the collection, return "Not in list"
    GetProvince_Collection = "Not in list"
End Function



Sub importRainfall()
    Dim myArray As Variant
    Dim rng As Range

    Select Case UCase(Range("T6").value)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
    End Select

    Dim indexString As String
    indexString = "data_" & UCase(Range("T6").value)

    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "기상청"
End Sub

'
' 2025/03/02 충남지역 버튼 추가
'
'

Sub importRainfall_button(ByVal AREA As String)
    Dim myArray As Variant
    Dim rng As Range

    Dim indexString As String
    indexString = "data_" & UCase(AREA)


    Select Case UCase(AREA)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
            
        Case "BORYUNG"
            Range("S5").value = "충청도"
            Range("T5").value = "보령"
        
        Case "DAEJEON"
            Range("S5").value = "충청도"
            Range("T5").value = "대전"
        
        Case "SEOSAN"
            Range("S5").value = "충청도"
            Range("T5").value = "서산"
        
        Case "BUYEO"
            Range("S5").value = "충청도"
            Range("T5").value = "부여"
        
        Case "CHEONAN"
            Range("S5").value = "충청도"
            Range("T5").value = "천안"
        
        Case "CHEONGJU"
            Range("S5").value = "충청도"
            Range("T5").value = "청주"
        
        Case "GEUMSAN"
            Range("S5").value = "충청도"
            Range("T5").value = "금산"
             
        
        Case "SEOUL"
            Range("S5").value = "서울경기"
            Range("T5").value = "서울"
            
        Case "SUWON"
            Range("S5").value = "서울경기"
            Range("T5").value = "수원"
            
        Case "INCHEON"
            Range("S5").value = "서울경기"
            Range("T5").value = "인천"
            
    End Select



    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "기상청"
End Sub




Function GetProvince_Case(city As String) As String
    Select Case city
        ' 충청도
        Case "보은", "제천", "청주", "추풍령", "대전", "세종", "금산", "보령", "부여", "서산", "천안", "홍성"
            GetProvince_Case = "충청도"
        ' 서울경기
        Case "관악산", "서울", "강화", "백령도", "인천", "동두천", "수원", "양평", "이천", "파주"
            GetProvince_Case = "서울경기"
        ' 강원도
        Case "강릉", "대관령", "동해", "북강릉", "북춘천", "삼척", "속초", "영월", "원주", "인제", "정선군", "철원", "춘천", "태백", "홍천"
            GetProvince_Case = "강원도"
        ' 전라도
        Case "광주", "고창", "고창군", "군산", "남원", "부안", "순창군", "임실", "장수", "전주", "정읍", "강진군", "고흥", "광양시", "목포", "무안", "보성군", "순천", "여수", "영광군", "완도", "장흥", "주암", "진도", "첨철산", "진도군", "해남", "흑산도"
            GetProvince_Case = "전라도"
        ' 경상도
        Case "대구", "대구(기)", "울산", "부산", "경주시", "구미", "문경", "봉화", "상주", "안동", "영덕", "영주", "영천", "울릉도", "울진", "의성", "청송군", "포항", "거제", "거창", "김해시", "남해", "밀양", "북창원", "산청", "양산시", "의령군", "진주", "창원", "통영", "함양군", "합천"
            GetProvince_Case = "경상도"
        ' 제주도
        Case "고산", "서귀포", "성산", "성산2", "성산포"
            GetProvince_Case = "제주도"
        ' Default case
        Case Else
            GetProvince_Case = "Not in list"
    End Select
End Function



Sub ResetWeatherData(ByVal AREA As String)

    Dim Province As String
    
    Sheets("All").Activate
'    Range("S5") = "충청도"
'    Range("T5") = "청주"
    
    Province = GetProvince_Case(AREA)

    If CheckSubstring(Province, "Not in list") Then
        Popup_MessageBox (" Province is Not in list .... ")
        Exit Sub
    End If

    Range("S5") = Province
    Range("T5") = AREA
    
    
    Popup_MessageBox ("Clear Contents")
    Range("b5:n34").ClearContents
    
    Popup_MessageBox (" Load 30 year Weather Data ")
    Call modProvince.importRainfall
    

End Sub


Sub test()

    Call ResetWeatherData("대전")

End Sub

