Attribute VB_Name = "BaseData_modEffectiveRadius"
Option Explicit

'0 : skin factor
'1 : Re1
'2 : Re2
'3 : Re3

Public Enum ER_VALUE
    erRE0 = 0
    erRE1 = 1
    erRE2 = 2
    erRE3 = 3
End Enum

Function GetER_Mode(ByVal WB_NAME As String) As Integer
    Dim er, R       As String
    
    ' er = Range("h10").value
    er = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("h10").value
    'MsgBox er
    R = Mid(er, 5, 1)
    
    If R = "F" Then
        GetER_Mode = 0
    Else
        GetER_Mode = val(R)
    End If
End Function



Function GetEffectiveRadius(ByVal WB_NAME As String) As Double
    Dim i, er As Integer
    
    If Not IsWorkBookOpen(WB_NAME) Then
        MsgBox "Please open the yangsoo data ! " & WB_NAME
        Exit Function
    End If
    
    er = GetER_Mode(WB_NAME)
    'Worksheets("SkinFactor").Range("k8").value  - 경험식 1번 (RE1)
    'Worksheets("SkinFactor").Range("k9").value  - 경험식 2번 (RE2)
    'Worksheets("SkinFactor").Range("k10").value  - 경험식 3번 (RE3)
    
    
    Select Case er
        Case erRE1
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k8").value
        Case erRE2
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k9").value
        Case erRE3
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k10").value
        Case Else
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("C8").value
    End Select

End Function


Function GetER_ModeFX(ByVal well_no As Integer) As Integer
    Dim er, R  As String
    Dim wsYangSoo As Worksheet
    
    Set wsYangSoo = Worksheets("YangSoo")
    
    ' ak : ER Mode
    er = wsYangSoo.Cells(4 + well_no, "ak").value
    
    'MsgBox er
    R = Mid(er, 5, 1)
    
    If R = "F" Then
        GetER_ModeFX = 0
    Else
        GetER_ModeFX = val(R)
    End If
End Function



Function GetEffectiveRadiusFromFX(ByVal well_no As Integer) As Double
    Dim i, er As Integer
    Dim wsYangSoo As Worksheet
    
    Set wsYangSoo = Worksheets("YangSoo")
    
    er = GetER_ModeFX(well_no)
    i = well_no
    
    'Worksheets("SkinFactor").Range("k8").value  - 경험식 1번 (RE1)
    'Worksheets("SkinFactor").Range("k9").value  - 경험식 2번 (RE2)
    'Worksheets("SkinFactor").Range("k10").value  - 경험식 3번 (RE3)
    
    Select Case er
        Case erRE1
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "AL").value
        Case erRE2
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "AM").value
        Case erRE3
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "AN").value
        Case Else
            GetEffectiveRadiusFromFX = wsYangSoo.Cells(4 + i, "Z").value
    End Select

End Function

