Sub ��ũ��1()
'
' ��ũ��1 ��ũ��
'

'
    Range("C14:P14").Select
    With Selection.Interior
        .pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .themeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("H17").Select
End Sub
Sub ��ũ��2()
'
' ��ũ��2 ��ũ��
'

'
    Range("C5:P17").Select
    With Selection.Interior
        .pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
