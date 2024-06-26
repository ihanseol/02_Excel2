Attribute VB_Name = "BaseData_DataJOjung"
Option Explicit

Dim ColorValue(1 To 20) As Long

Public Sub InitialSetColorValue()
    ColorValue(1) = RGB(192, 0, 0)
    ColorValue(2) = RGB(255, 0, 0)
    ColorValue(3) = RGB(255, 192, 0)
    ColorValue(4) = RGB(255, 255, 0)
    ColorValue(5) = RGB(146, 208, 80)
    ColorValue(6) = RGB(0, 176, 80)
    ColorValue(7) = RGB(0, 176, 240)
    ColorValue(8) = RGB(0, 112, 192)
    ColorValue(9) = RGB(0, 32, 96)
    ColorValue(10) = RGB(112, 48, 160)
    
    ColorValue(11) = RGB(192 + 10, 10, 0)
    ColorValue(12) = RGB(255, 0 + 10, 0)
    ColorValue(13) = RGB(255, 192 + 10, 0)
    ColorValue(14) = RGB(255, 255, 10)
    ColorValue(15) = RGB(146 + 10, 208 + 10, 80 + 10)
    ColorValue(16) = RGB(0 + 10, 176 + 10, 80)
    ColorValue(17) = RGB(0 + 10, 176 + 10, 240 + 10)
    ColorValue(18) = RGB(0 + 10, 112 + 10, 192)
    ColorValue(19) = RGB(0 + 10, 32 + 10, 96)
    ColorValue(20) = RGB(112, 48 + 10, 160 + 10)
End Sub

Private Sub initialize_wellstyle()
    
    Dim rng, cell As Range

    Set rng = Range("C3:C22")
    Range("C3:C22").Select
    
    Selection.NumberFormat = "General"
        
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' 2024/6/15
    For Each cell In rng
        SetFontAndInteriorColorBasedOnBackground cell
    Next cell

    
    Range("E19:G19").Select
    With Selection.Font
        .name = "∏º¿∫ ∞ÌµÒ"
        .Size = 12
        .themeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("E21:G21").Select
    With Selection.Font
        .name = "∏º¿∫ ∞ÌµÒ"
        .Size = 12
        .themeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B25:K29").Select
    With Selection.Font
        .name = "∏º¿∫ ∞ÌµÒ"
        .Size = 11
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("d23").Select
End Sub


' 2024/6/15
Sub test_SetFontAndInteriorColorBasedOnBackground()

    Dim cell, rng As Range
    
    Set rng = ActiveSheet.Range("c7, c8, c9")

    For Each cell In rng
        SetFontAndInteriorColorBasedOnBackground cell
    Next cell

End Sub


' 2024/6/15
Function GetBackgroundColor(ByVal cell As Range) As Long
    GetBackgroundColor = cell.Interior.color
End Function


' Subroutine to set the font and interior color based on the background color
Sub SetFontAndInteriorColorBasedOnBackground(ByVal cell As Range)
    Dim bgColor As Long
    
    ' Get the background color of the cell
    bgColor = GetBackgroundColor(cell)
    
    ' Determine if the background color is dark
    If IsDarkColor(bgColor) Then
        With cell.Font
            .name = "∏º¿∫ ∞ÌµÒ"
            .Size = 10
            .themeColor = xlThemeColorDark1 ' Light font color for dark background
            .ThemeFont = xlThemeFontNone
        End With
    Else
        With cell.Font
            .name = "∏º¿∫ ∞ÌµÒ"
            .Size = 10
            .themeColor = xlThemeColorLight1 ' Dark font color for light background
            .ThemeFont = xlThemeFontNone
        End With
    End If
End Sub

' Function to determine if a color is dark
Function IsDarkColor(color As Long) As Boolean
    Dim R As Long, G As Long, B As Long
    R = (color Mod 256)
    G = ((color \ 256) Mod 256)
    B = ((color \ 65536) Mod 256)
    
    ' Calculate brightness (perceived luminance)
    ' Using the formula: 0.299*R + 0.587*G + 0.114*B
    If (0.299 * R + 0.587 * G + 0.114 * B) < 128 Then
        IsDarkColor = True
    Else
        IsDarkColor = False
    End If
End Function


' 2024/6/15
Sub DetermineThemeColor()
    Dim ws As Worksheet
    Dim cell As Range
    Dim themeColor As Long
    
    ' Set your worksheet and cell
    Set ws = ActiveSheet
    Set cell = ws.Range("c8")
    
    ' Get the theme color if it exists
    On Error Resume Next
    themeColor = cell.Interior.themeColor
    On Error GoTo 0
    
    ' Check if the theme color is valid
    If themeColor <> xlColorIndexNone Then
        MsgBox "The theme color of the cell is: " & ThemeColorName(themeColor)
    Else
        MsgBox "The cell does not have a theme color."
    End If
End Sub


' 2024/6/15
Function ThemeColorName(themeColor As Long) As String
    Select Case themeColor
        Case xlThemeColorDark1
            ThemeColorName = "Dark1"
        Case xlThemeColorLight1
            ThemeColorName = "Light1"
        Case xlThemeColorDark2
            ThemeColorName = "Dark2"
        Case xlThemeColorLight2
            ThemeColorName = "Light2"
        Case xlThemeColorAccent1
            ThemeColorName = "Accent1"
        Case xlThemeColorAccent2
            ThemeColorName = "Accent2"
        Case xlThemeColorAccent3
            ThemeColorName = "Accent3"
        Case xlThemeColorAccent4
            ThemeColorName = "Accent4"
        Case xlThemeColorAccent5
            ThemeColorName = "Accent5"
        Case xlThemeColorAccent6
            ThemeColorName = "Accent6"
        Case xlThemeColorHyperlink
            ThemeColorName = "Hyperlink"
        Case xlThemeColorFollowedHyperlink
            ThemeColorName = "Followed Hyperlink"
        Case Else
            ThemeColorName = "Unknown Theme Color"
    End Select
End Function



Private Sub change_font_size()
    Range("J25").Select
    Selection.Font.Size = 10
    Range("F26").Select
    Selection.Font.Size = 10
End Sub

Public Sub make_wellstyle()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    Call TurnOffStuff
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        Call initialize_wellstyle
        Call change_font_size
    Next i
    
    Call TurnOnStuff
    
End Sub

Private Sub JojungData(ByVal nsheet As Integer)
    Dim nselect     As String
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate
    
    nsheet = nsheet + 3
    '=Well!D7
    nselect = Mid(Range("c2").formula, 8)
    
    'Debug.Print Mid(Range("c2").Formula, 8) & ":" & nselect
    
    Selection.Replace What:=nselect, Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    ' minhwasoo 2023/10/13
    ' Range("E21").Select
    ' Range("E21").formula = "=Well!" & Cells(nsheet, "I").Address
End Sub

Private Sub SetMyTabColor(ByVal index As Integer)
    If Sheets("Well").SingleColor.value Then
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .color = 192
            .TintAndShade = 0
        End With
    Else
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .color = ColorValue(index)
            .TintAndShade = 0
        End With
    End If
End Sub

'∞¢∞¢¿« Ω¨∆Æ∏¶ º¯»∏«œ∏Èº≠, ºø¿« ¬¸¡∂∞™¿ª ê¨√ﬂæÓ¡ÿ¥Ÿ.
'
Public Sub JojungSheetData()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Cells(i + 3, "A").value = "W" & i
    Next i
    
    For i = 1 To n_sheets
        Sheets(CStr(i)).Activate
        Range("B26").value = "W-" & i
        
        Call JojungData(i)
        Call SetMyTabColor(i)
    Next i
End Sub

